import argparse
import re
import socket
import struct
import subprocess
import sys
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from pywinauto import Application, Desktop, mouse

from freecell_solver import (
    CARD_PORT,
    FRAME_SIZE,
    FreeCellSolver,
    decode_card,
    decode_foundations,
)


SUIT_INDEX = {"S": 0, "D": 1, "H": 2, "C": 3}
MOVE_RE = re.compile(r"^Move (?P<card>\S+) from (?P<src>.+?) to (?P<dst>.+)$")


@dataclass
class BoardState:
    columns: list[list[str]]
    freecells: list[Optional[str]]
    foundations: list[int]


@dataclass
class Move:
    raw: str
    card: str
    src_kind: str
    src_index: Optional[int]
    dst_kind: str
    dst_index: Optional[int]


class AutoSolver:
    def __init__(self, running: bool):
        self.base_dir = Path(__file__).resolve().parent
        self.exe_path = self.base_dir / "Group_Project_T2" / "freecell.exe"
        self.find_cards_path = self.base_dir / "find_cards.exe"
        self.solution_path = self.base_dir / "solution.txt"

        if running:
            self.app = self._connect_to_game()
        else:
            self.app = Application(backend="win32").start(str(self.exe_path))

        self.window = self.app.window(title_re="FreeCell.*")
        self.window.wait("visible ready", timeout=15)
        self.window.set_focus()

    def _connect_to_game(self) -> Application:
        app = Application(backend="win32")
        try:
            app.connect(path=str(self.exe_path))
        except Exception:
            app.connect(title_re="FreeCell.*")
        return app

    def _refresh_window(self) -> None:
        self.window = self.app.window(title_re="FreeCell.*")
        self.window.wait("visible ready", timeout=15)

    def _process_id(self) -> int:
        return int(self.window.element_info.process_id)

    def getControlIdentifiers(self):
        return self.window.print_control_identifiers()

    def startGame(self) -> None:
        self.window.menu_select("Game->New Game")
        self._refresh_window()
        self.window.set_focus()
        time.sleep(0.5)

    def quit(self) -> None:
        self.window.type_keys("%{F4}")
        time.sleep(0.25)

        for candidate in Desktop(backend="win32").windows(process=self._process_id()):
            if candidate.handle == self.window.handle:
                continue
            dialog = self.app.window(handle=candidate.handle)
            if dialog.exists(timeout=1):
                for child in dialog.children():
                    if child.friendly_class_name() == "Button" and child.window_text().lower() == "yes":
                        child.click()
                        return

    def solve_current_game(self) -> tuple[BoardState, list[str]]:
        board_state = self._capture_board_snapshot()
        solver = FreeCellSolver(
            board_state.columns,
            [card for card in board_state.freecells if card],
            tuple(board_state.foundations),
        )

        print("Solving...")
        solution = solver.solve()
        if not solution:
            raise RuntimeError("Solver did not find a solution for the current game.")

        self._write_solution(solution)
        return board_state, solution

    def _capture_board_snapshot(self) -> BoardState:
        if not self.find_cards_path.exists():
            raise FileNotFoundError(f"Missing card reader: {self.find_cards_path}")

        payload_holder: dict[str, bytes] = {}
        error_holder: dict[str, BaseException] = {}
        ready = threading.Event()

        def server() -> None:
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as listener:
                    listener.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                    listener.bind(("127.0.0.1", CARD_PORT))
                    listener.listen(1)
                    ready.set()
                    conn, _ = listener.accept()
                    with conn:
                        payload_holder["data"] = self._read_exact(conn, FRAME_SIZE)
            except BaseException as exc:  # pragma: no cover - surfaced to caller
                error_holder["error"] = exc
                ready.set()

        thread = threading.Thread(target=server, daemon=True)
        thread.start()

        if not ready.wait(timeout=5):
            raise TimeoutError("Timed out while waiting for the local card socket server.")

        if "error" in error_holder:
            raise RuntimeError("Failed to start the local card socket server.") from error_holder["error"]

        proc = subprocess.run(
            [str(self.find_cards_path)],
            cwd=self.base_dir,
            capture_output=True,
            text=True,
            timeout=15,
            check=False,
        )
        if proc.returncode != 0:
            raise RuntimeError(
                "find_cards.exe failed.\n"
                f"stdout:\n{proc.stdout}\n"
                f"stderr:\n{proc.stderr}"
            )

        thread.join(timeout=10)
        if thread.is_alive():
            raise TimeoutError("Timed out while waiting for card data from find_cards.exe.")
        if "error" in error_holder:
            raise RuntimeError("Card socket server failed while reading the board.") from error_holder["error"]
        if "data" not in payload_holder:
            raise RuntimeError("No board payload was received from find_cards.exe.")

        return self._decode_snapshot(payload_holder["data"])

    @staticmethod
    def _read_exact(sock: socket.socket, size: int) -> bytes:
        data = b""
        while len(data) < size:
            chunk = sock.recv(size - len(data))
            if not chunk:
                raise RuntimeError("Socket closed before the board payload was fully received.")
            data += chunk
        return data

    @staticmethod
    def _decode_snapshot(payload: bytes) -> BoardState:
        ints = struct.unpack("<193i", payload)

        cards = ints[:189]
        foundations = list(decode_foundations(ints[189:193]))
        freecells = [decode_card(cards[i]) for i in range(4)]

        columns: list[list[str]] = [[] for _ in range(8)]
        for col_idx in range(8):
            for offset in range(21):
                value = cards[(col_idx + 1) * 21 + offset]
                card = decode_card(value)
                if not card:
                    break
                columns[col_idx].append(card)

        return BoardState(columns=columns, freecells=freecells, foundations=foundations)

    def _write_solution(self, solution: list[str]) -> None:
        with self.solution_path.open("w", encoding="utf-8") as handle:
            for index, move in enumerate(solution, start=1):
                handle.write(f"{index}. {move}\n")

    def play_solution(self, board_state: BoardState, solution: list[str], move_delay: float = 0.12) -> None:
        self._bring_to_front()

        for raw_move in solution:
            move = self._parse_move(raw_move)
            self._execute_move(board_state, move)
            self._handle_single_card_popup()
            time.sleep(move_delay)

    def _bring_to_front(self) -> None:
        self._refresh_window()
        self.window.restore()
        self.window.set_focus()
        time.sleep(0.2)

    def _parse_move(self, raw_move: str) -> Move:
        text = raw_move.strip()
        match = MOVE_RE.match(text)
        if not match:
            raise ValueError(f"Unrecognized move format: {raw_move}")

        src_kind, src_index = self._parse_place(match.group("src"))
        dst_kind, dst_index = self._parse_place(match.group("dst"))
        return Move(
            raw=text,
            card=match.group("card"),
            src_kind=src_kind,
            src_index=src_index,
            dst_kind=dst_kind,
            dst_index=dst_index,
        )

    @staticmethod
    def _parse_place(place: str) -> tuple[str, Optional[int]]:
        place = place.strip()
        if place == "FreeCell":
            return "freecell", None
        if place == "Foundation":
            return "foundation", None
        match = re.fullmatch(r"(?:empty )?Col (\d+)", place)
        if match:
            return "column", int(match.group(1)) - 1
        raise ValueError(f"Unrecognized location: {place}")

    def _execute_move(self, board_state: BoardState, move: Move) -> None:
        src_point = self._source_point(board_state, move)
        dst_point = self._destination_point(board_state, move)

        mouse.press(button="left", coords=src_point)
        time.sleep(0.04)
        mouse.move(coords=dst_point)
        time.sleep(0.04)
        mouse.release(button="left", coords=dst_point)

        self._apply_move_to_state(board_state, move)

    def _source_point(self, board_state: BoardState, move: Move) -> tuple[int, int]:
        if move.src_kind == "column":
            assert move.src_index is not None
            depth = len(board_state.columns[move.src_index]) - 1
            if depth < 0:
                raise RuntimeError(f"Column {move.src_index + 1} is empty before move: {move.raw}")
            return self._column_card_point(move.src_index, depth)

        if move.src_kind == "freecell":
            slot = self._find_freecell_slot(board_state, move.card)
            return self._slot_point(slot)

        raise RuntimeError(f"Unsupported move source: {move.raw}")

    def _destination_point(self, board_state: BoardState, move: Move) -> tuple[int, int]:
        if move.dst_kind == "column":
            assert move.dst_index is not None
            depth = len(board_state.columns[move.dst_index]) - 1
            return self._empty_column_point(move.dst_index) if depth < 0 else self._column_card_point(move.dst_index, depth)

        if move.dst_kind == "freecell":
            slot = self._first_empty_freecell(board_state)
            return self._slot_point(slot)

        if move.dst_kind == "foundation":
            return self._foundation_point(SUIT_INDEX[move.card[-1]])

        raise RuntimeError(f"Unsupported move destination: {move.raw}")

    def _apply_move_to_state(self, board_state: BoardState, move: Move) -> None:
        if move.src_kind == "column":
            assert move.src_index is not None
            moved_card = board_state.columns[move.src_index].pop()
        elif move.src_kind == "freecell":
            slot = self._find_freecell_slot(board_state, move.card)
            moved_card = board_state.freecells[slot]
            board_state.freecells[slot] = None
        else:
            raise RuntimeError(f"Unsupported move source: {move.raw}")

        if moved_card != move.card:
            raise RuntimeError(f"State mismatch while replaying move: {move.raw}")

        if move.dst_kind == "column":
            assert move.dst_index is not None
            board_state.columns[move.dst_index].append(move.card)
            return

        if move.dst_kind == "freecell":
            slot = self._first_empty_freecell(board_state)
            board_state.freecells[slot] = move.card
            return

        if move.dst_kind == "foundation":
            suit_index = SUIT_INDEX[move.card[-1]]
            board_state.foundations[suit_index] += 1
            return

        raise RuntimeError(f"Unsupported move destination: {move.raw}")

    @staticmethod
    def _find_freecell_slot(board_state: BoardState, card: str) -> int:
        for index, value in enumerate(board_state.freecells):
            if value == card:
                return index
        raise RuntimeError(f"Card {card} was not found in any freecell.")

    @staticmethod
    def _first_empty_freecell(board_state: BoardState) -> int:
        for index, value in enumerate(board_state.freecells):
            if value is None:
                return index
        raise RuntimeError("No empty freecell slot is available.")

    def _handle_single_card_popup(self) -> None:
        deadline = time.time() + 1.5
        while time.time() < deadline:
            dialog = self._find_choice_dialog()
            if not dialog:
                return

            buttons = [child for child in dialog.children() if child.friendly_class_name() == "Button"]
            preferred = self._pick_single_card_button(buttons)
            if preferred is None:
                return

            preferred.click()
            time.sleep(0.15)

    def _find_choice_dialog(self):
        self._refresh_window()
        for candidate in Desktop(backend="win32").windows(process=self._process_id()):
            if candidate.handle == self.window.handle or not candidate.is_visible():
                continue
            dialog = self.app.window(handle=candidate.handle)
            if dialog.exists(timeout=0.2):
                return dialog
        return None

    @staticmethod
    def _pick_single_card_button(buttons):
        if not buttons:
            return None

        patterns = (
            re.compile(r"one", re.IGNORECASE),
            re.compile(r"single", re.IGNORECASE),
            re.compile(r"card", re.IGNORECASE),
            re.compile(r"just", re.IGNORECASE),
        )
        for pattern in patterns:
            for button in buttons:
                if pattern.search(button.window_text()):
                    return button

        if len(buttons) == 2:
            return min(buttons, key=lambda btn: btn.rectangle().left)

        return buttons[0]

    def _slot_point(self, slot_index: int) -> tuple[int, int]:
        metrics = self._board_metrics()
        x = metrics["left_margin"] + slot_index * metrics["slot_pitch"] + (metrics["card_width"] / 2.0)
        y = metrics["top_slots_top"] + (metrics["card_height"] / 2.0)
        return self._client_to_screen(x, y)

    def _foundation_point(self, foundation_index: int) -> tuple[int, int]:
        return self._slot_point(4 + foundation_index)

    def _empty_column_point(self, column_index: int) -> tuple[int, int]:
        metrics = self._board_metrics()
        x = metrics["left_margin"] + column_index * metrics["slot_pitch"] + (metrics["card_width"] / 2.0)
        y = metrics["tableau_top"] + (metrics["card_height"] / 2.0)
        return self._client_to_screen(x, y)

    def _column_card_point(self, column_index: int, depth: int) -> tuple[int, int]:
        metrics = self._board_metrics()
        x = metrics["left_margin"] + column_index * metrics["slot_pitch"] + (metrics["card_width"] / 2.0)
        y = metrics["tableau_top"] + depth * metrics["card_step"] + min(metrics["card_height"] * 0.35, 34.0)
        return self._client_to_screen(x, y)

    def _board_metrics(self) -> dict[str, float]:
        rect = self.window.client_rect()
        width = float(rect.right - rect.left)

        left_margin = max(8.0, round(width / 80.0))
        slot_gap = left_margin
        card_width = (width - 2.0 * left_margin - 7.0 * slot_gap) / 8.0
        card_height = card_width * (96.0 / 71.0)
        menu_height = max(20.0, round(card_height * 0.22))
        top_slots_top = menu_height + left_margin
        tableau_top = top_slots_top + card_height + max(18.0, round(card_height * 0.15))
        card_step = max(18.0, round(card_height * 0.20))

        return {
            "left_margin": left_margin,
            "slot_pitch": card_width + slot_gap,
            "card_width": card_width,
            "card_height": card_height,
            "top_slots_top": top_slots_top,
            "tableau_top": tableau_top,
            "card_step": card_step,
        }

    def _client_to_screen(self, x: float, y: float) -> tuple[int, int]:
        client_x, client_y = self.window.client_to_screen((int(round(x)), int(round(y))))
        return int(client_x), int(client_y)


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="AutoSolver",
        description="Starts or binds to Windows FreeCell, captures the board, solves it, and replays the moves.",
    )
    parser.add_argument(
        "-r",
        "--running",
        action="store_true",
        help="Bind to an already-running FreeCell window instead of launching a new one.",
    )
    parser.add_argument(
        "--leave-open",
        action="store_true",
        help="Do not close the game window after the automated solution finishes.",
    )
    parser.add_argument(
        "--move-delay",
        type=float,
        default=0.12,
        help="Delay in seconds between replayed moves.",
    )
    args = parser.parse_args()

    solver = AutoSolver(args.running)
    solver.startGame()

    board_state, solution = solver.solve_current_game()
    print(f"Wrote {len(solution)} moves to {solver.solution_path}")

    solver.play_solution(board_state, solution, move_delay=args.move_delay)

    if not args.leave_open:
        time.sleep(0.5)
        solver.quit()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"AutoSolver failed: {exc}", file=sys.stderr)
        raise
