import argparse
import ctypes
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

from pywinauto import Application, Desktop

from freecell_solver import (
    CARD_PORT,
    FRAME_SIZE,
    FreeCellSolver,
    decode_card,
    decode_foundations,
)


SUIT_INDEX = {"S": 0, "D": 1, "H": 2, "C": 3}
MOVE_RE = re.compile(r"^Move (?P<card>\S+) from (?P<src>.+?) to (?P<dst>.+)$")

# Toggle this back to False if you want AutoSolver to replay foundation moves explicitly.
SKIP_FOUNDATION_MOVES = False
# Keep invalid-move dialogs visible in the logs and fail fast after dismissing them.
INVALID_MOVE_IS_FATAL = True

WM_MOUSEMOVE = 0x0200
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
MK_LBUTTON = 0x0001


@dataclass
class BoardState:
    columns: list[list[str]]
    freecells: list[Optional[str]]
    foundations: list[int]
    foundation_slots: list[Optional[str]]


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
        self.main_window_handle: Optional[int] = None
        self.foundation_slot_by_suit: dict[str, int] = {}

        if running:
            self.app = self._connect_to_game()
        else:
            self.app = Application(backend="win32").start(str(self.exe_path))

        self.window = self._resolve_main_window()
        self.window.set_focus()

    def _connect_to_game(self) -> Application:
        candidates = []
        for candidate in Desktop(backend="win32").windows():
            if not candidate.is_visible():
                continue
            if not re.match(r"FreeCell.*", candidate.window_text() or ""):
                continue
            candidates.append(candidate)

        if not candidates:
            raise RuntimeError("Could not find a running FreeCell window to attach to.")

        candidates.sort(key=lambda window: (window.process_id(), window.handle))
        chosen = candidates[0]
        print(
            f"[debug] Attaching to running FreeCell process={chosen.process_id()} "
            f"handle={chosen.handle} title={chosen.window_text()!r} among {len(candidates)} candidate(s)"
        )

        app = Application(backend="win32")
        app.connect(process=chosen.process_id())
        self.main_window_handle = chosen.handle
        return app

    def _refresh_window(self) -> None:
        self.window = self._resolve_main_window()
        self.window.wait("visible ready", timeout=15)

    def _resolve_main_window(self):
        if self.main_window_handle is not None:
            window = self.app.window(handle=self.main_window_handle)
            if window.exists(timeout=1):
                return window

        candidates = []
        for candidate in Desktop(backend="win32").windows(process=self.app.process):
            if not candidate.is_visible():
                continue
            if not re.match(r"FreeCell.*", candidate.window_text() or ""):
                continue
            candidates.append(candidate)

        if not candidates:
            raise RuntimeError("Could not find the main FreeCell window.")

        candidates.sort(key=lambda window: (window.handle, window.window_text()))
        chosen = candidates[0]
        self.main_window_handle = chosen.handle
        print(
            f"[debug] Selected FreeCell window handle={chosen.handle} "
            f"title={chosen.window_text()!r} among {len(candidates)} candidate(s)"
        )
        return self.app.window(handle=chosen.handle)

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
        raw_foundations = ints[189:193]
        foundations = list(decode_foundations(raw_foundations))
        foundation_slots = [decode_card(value) for value in raw_foundations]
        freecells = [decode_card(cards[i]) for i in range(4)]

        columns: list[list[str]] = [[] for _ in range(8)]
        for col_idx in range(8):
            for offset in range(21):
                value = cards[(col_idx + 1) * 21 + offset]
                card = decode_card(value)
                if not card:
                    break
                columns[col_idx].append(card)

        return BoardState(
            columns=columns,
            freecells=freecells,
            foundations=foundations,
            foundation_slots=foundation_slots,
        )

    def _write_solution(self, solution: list[str]) -> None:
        with self.solution_path.open("w", encoding="utf-8") as handle:
            for index, move in enumerate(solution, start=1):
                handle.write(f"{index}. {move}\n")

    def play_solution(self, board_state: BoardState, solution: list[str], move_delay: float = 0.12) -> None:
        self._bring_to_front()

        for move_number, raw_move in enumerate(solution, start=1):
            board_state = self._capture_board_snapshot()
            self._sync_foundation_slot_map(board_state)
            move = self._parse_move(raw_move)
            source_mismatch = self._source_mismatch(board_state, move)
            if source_mismatch:
                print(f"[debug] Skipping move {move_number}: {move.raw} | {source_mismatch}")
                continue
            self._ensure_live_foundation_mapping(board_state, move)
            if SKIP_FOUNDATION_MOVES and move.dst_kind == "foundation":
                print(f"[debug] Skipping move {move_number}: {move.raw}")
                self._apply_move_to_state(board_state, move)
                continue
            self._execute_move(board_state, move_number, move)
            self._handle_single_card_popup()
            invalid_message = self._handle_invalid_move_popup()
            if invalid_message and INVALID_MOVE_IS_FATAL:
                raise RuntimeError(f"FreeCell rejected move {move_number}: {move.raw} | dialog={invalid_message}")
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

    def _execute_move(self, board_state: BoardState, move_number: int, move: Move) -> None:
        src_point = self._source_point(board_state, move)
        dst_point = self._destination_point(board_state, move)
        src_label = self._describe_source(board_state, move)
        dst_label = self._describe_destination(board_state, move)
        print(
            f"[debug] Move {move_number}: {move.raw} | "
            f"source={src_label} destination={dst_label} "
            f"src={src_point} dst={dst_point} "
            f"freecells={board_state.freecells} foundations={board_state.foundations}"
        )

        self._send_pick_and_place(src_point, dst_point)

        self._apply_move_to_state(board_state, move)
        print(
            f"[debug] Move {move_number} applied | "
            f"freecells={board_state.freecells} foundations={board_state.foundations}"
        )

    def _source_point(self, board_state: BoardState, move: Move) -> tuple[int, int]:
        if move.src_kind == "column":
            assert move.src_index is not None
            depth = len(board_state.columns[move.src_index]) - 1
            if depth < 0:
                raise RuntimeError(f"Column {move.src_index + 1} is empty before move: {move.raw}")
            return self._source_column_point(move.src_index, depth)

        if move.src_kind == "freecell":
            slot = self._find_freecell_slot(board_state, move.card)
            return self._freecell_source_point(slot)

        raise RuntimeError(f"Unsupported move source: {move.raw}")

    def _destination_point(self, board_state: BoardState, move: Move) -> tuple[int, int]:
        if move.dst_kind == "column":
            assert move.dst_index is not None
            depth = len(board_state.columns[move.dst_index]) - 1
            return self._empty_column_point(move.dst_index) if depth < 0 else self._target_column_point(move.dst_index, depth)

        if move.dst_kind == "freecell":
            slot = self._first_empty_freecell(board_state)
            return self._freecell_target_point(slot)

        if move.dst_kind == "foundation":
            return self._foundation_point(self._foundation_slot_for_suit(board_state, move.card[-1]))

        raise RuntimeError(f"Unsupported move destination: {move.raw}")

    def _describe_source(self, board_state: BoardState, move: Move) -> str:
        if move.src_kind == "column":
            assert move.src_index is not None
            return f"Column {move.src_index + 1}"

        if move.src_kind == "freecell":
            slot = self._find_freecell_slot(board_state, move.card)
            return f"FreeCell {slot + 1}"

        return move.src_kind

    def _describe_destination(self, board_state: BoardState, move: Move) -> str:
        if move.dst_kind == "column":
            assert move.dst_index is not None
            return f"Column {move.dst_index + 1}"

        if move.dst_kind == "freecell":
            slot = self._first_empty_freecell(board_state)
            return f"FreeCell {slot + 1}"

        if move.dst_kind == "foundation":
            suit = move.card[-1]
            foundation_slot = self._foundation_slot_for_suit(board_state, suit) + 1
            suit_name = {
                "S": "Spades",
                "D": "Diamonds",
                "H": "Hearts",
                "C": "Clubs",
            }[suit]
            return f"Foundation {foundation_slot} ({suit_name})"

        return move.dst_kind

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
            slot_index = self._foundation_slot_for_suit(board_state, move.card[-1])
            board_state.foundation_slots[slot_index] = move.card
            return

        raise RuntimeError(f"Unsupported move destination: {move.raw}")

    def _source_mismatch(self, board_state: BoardState, move: Move) -> Optional[str]:
        if move.src_kind == "column":
            assert move.src_index is not None
            column = board_state.columns[move.src_index]
            if not column:
                return f"source column {move.src_index + 1} is empty"
            if column[-1] != move.card:
                return (
                    f"expected {move.card} on top of column {move.src_index + 1}, "
                    f"found {column[-1]}"
                )
            return None

        if move.src_kind == "freecell":
            try:
                self._find_freecell_slot(board_state, move.card)
            except RuntimeError:
                return f"expected {move.card} in a freecell, but it is no longer there"
            return None

        return f"unsupported move source: {move.raw}"

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

    @staticmethod
    def _first_empty_foundation_slot(board_state: BoardState) -> int:
        for index, value in enumerate(board_state.foundation_slots):
            if value is None:
                return index
        raise RuntimeError("No empty foundation slot is available.")

    def _sync_foundation_slot_map(self, board_state: BoardState) -> None:
        for index, value in enumerate(board_state.foundation_slots):
            if value:
                suit = value[-1]
                mapped_index = self.foundation_slot_by_suit.get(suit)
                if mapped_index is not None and mapped_index != index:
                    print(
                        f"[debug] Foundation slot remap for suit {suit}: "
                        f"{mapped_index + 1} -> {index + 1}"
                    )
                self.foundation_slot_by_suit[suit] = index

    def _foundation_slot_for_suit(self, board_state: BoardState, suit: str) -> int:
        mapped_index = self.foundation_slot_by_suit.get(suit)
        if mapped_index is not None:
            return mapped_index

        self._sync_foundation_slot_map(board_state)
        mapped_index = self.foundation_slot_by_suit.get(suit)
        if mapped_index is not None:
            return mapped_index

        mapped_index = self._first_empty_foundation_slot(board_state)
        self.foundation_slot_by_suit[suit] = mapped_index
        print(f"[debug] Assigned suit {suit} to foundation slot {mapped_index + 1}")
        return mapped_index

    def _ensure_live_foundation_mapping(self, board_state: BoardState, move: Move) -> None:
        if move.dst_kind != "foundation":
            return

        suit = move.card[-1]
        if suit in self.foundation_slot_by_suit:
            return

        for index, value in enumerate(board_state.foundation_slots):
            if value and value[-1] == suit:
                self.foundation_slot_by_suit[suit] = index
                print(
                    f"[debug] Learned live foundation slot for suit {suit}: "
                    f"slot {index + 1} from {value}"
                )
                return


    def _handle_single_card_popup(self) -> None:
        deadline = time.time() + 1.5
        while time.time() < deadline:
            dialog = self._find_choice_dialog()
            if not dialog:
                return

            buttons = [child for child in dialog.children() if child.friendly_class_name() == "Button"]
            preferred = self._pick_single_card_button(buttons)
            if preferred is None:
                print("[debug] Choice dialog detected but no button could be selected")
                return

            print(
                f"[debug] Choice dialog handle={dialog.handle} title={dialog.window_text()!r} "
                f"buttons={[button.window_text() for button in buttons]} "
                f"selected={preferred.window_text()!r}"
            )
            preferred.click()
            time.sleep(0.15)

    def _handle_invalid_move_popup(self) -> Optional[str]:
        deadline = time.time() + 1.5
        while time.time() < deadline:
            dialog = self._find_invalid_move_dialog()
            if not dialog:
                return None

            buttons = [child for child in dialog.children() if child.friendly_class_name() == "Button"]
            ok_button = self._pick_ok_button(buttons)
            texts = [dialog.window_text()]
            texts.extend(child.window_text() for child in dialog.children())
            message = " | ".join(text for text in texts if text)
            if ok_button is None:
                print(
                    f"[debug] Invalid move dialog handle={dialog.handle} title={dialog.window_text()!r} "
                    "but no OK button was found"
                )
                return message

            print(
                f"[debug] Invalid move dialog handle={dialog.handle} title={dialog.window_text()!r} "
                f"buttons={[button.window_text() for button in buttons]} "
                f"selected={ok_button.window_text()!r}"
            )
            ok_button.click()
            time.sleep(0.15)
            return message
        return None

    def _find_choice_dialog(self):
        self._refresh_window()
        for candidate in Desktop(backend="win32").windows(process=self._process_id()):
            if candidate.handle == self.window.handle or not candidate.is_visible():
                continue
            dialog = self.app.window(handle=candidate.handle)
            if dialog.exists(timeout=0.2):
                return dialog
        return None

    def _find_invalid_move_dialog(self):
        self._refresh_window()
        for candidate in Desktop(backend="win32").windows(process=self._process_id()):
            if candidate.handle == self.window.handle or not candidate.is_visible():
                continue

            dialog = self.app.window(handle=candidate.handle)
            if not dialog.exists(timeout=0.2):
                continue

            texts = [dialog.window_text()]
            texts.extend(child.window_text() for child in dialog.children())
            normalized = " ".join(text for text in texts if text).lower()
            if "invalid move" in normalized:
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

    @staticmethod
    def _pick_ok_button(buttons):
        for button in buttons:
            if button.window_text().strip().lower() in {"ok", "&ok"}:
                return button
        if len(buttons) == 1:
            return buttons[0]
        return None

    def _slot_point(self, slot_index: int) -> tuple[int, int]:
        metrics = self._board_metrics()
        x = metrics["left_margin"] + slot_index * metrics["slot_pitch"] + (metrics["card_width"] / 2.0)
        y = metrics["top_slots_top"] + (metrics["card_height"] / 2.0)
        return self._client_point(x, y)

    def _freecell_source_point(self, slot_index: int) -> tuple[int, int]:
        metrics = self._board_metrics()
        x = metrics["left_margin"] + slot_index * metrics["slot_pitch"] + (metrics["card_width"] / 2.0)
        y = metrics["top_slots_top"] + min(metrics["card_height"] * 0.60, 52.0)
        return self._client_point(x, y)

    def _freecell_target_point(self, slot_index: int) -> tuple[int, int]:
        metrics = self._board_metrics()
        x = metrics["left_margin"] + slot_index * metrics["slot_pitch"] + (metrics["card_width"] / 2.0)
        y = metrics["top_slots_top"] + min(metrics["card_height"] * 0.30, 28.0)
        return self._client_point(x, y)

    def _foundation_point(self, foundation_index: int) -> tuple[int, int]:
        return self._slot_point(4 + foundation_index)

    def _empty_column_point(self, column_index: int) -> tuple[int, int]:
        metrics = self._board_metrics()
        x = metrics["left_margin"] + column_index * metrics["slot_pitch"] + (metrics["card_width"] / 2.0)
        y = metrics["tableau_top"] + (metrics["card_height"] / 2.0)
        return self._client_point(x, y)

    def _source_column_point(self, column_index: int, depth: int) -> tuple[int, int]:
        metrics = self._board_metrics()
        x = metrics["left_margin"] + column_index * metrics["slot_pitch"] + (metrics["card_width"] / 2.0)
        y = metrics["tableau_top"] + depth * metrics["card_step"] + min(metrics["card_height"] * 0.60, 52.0)
        return self._client_point(x, y)

    def _target_column_point(self, column_index: int, depth: int) -> tuple[int, int]:
        metrics = self._board_metrics()
        x = metrics["left_margin"] + column_index * metrics["slot_pitch"] + (metrics["card_width"] / 2.0)
        y = metrics["tableau_top"] + depth * metrics["card_step"] + max(8.0, min(metrics["card_step"] * 0.5, 18.0))
        return self._client_point(x, y)

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

    @staticmethod
    def _client_point(x: float, y: float) -> tuple[int, int]:
        return int(round(x)), int(round(y))

    @staticmethod
    def _make_lparam(point: tuple[int, int]) -> int:
        x, y = point
        return (y << 16) | (x & 0xFFFF)

    def _send_mouse_message(self, msg: int, wparam: int, point: tuple[int, int]) -> None:
        ctypes.windll.user32.SendMessageW(
            int(self.window.handle),
            msg,
            wparam,
            self._make_lparam(point),
        )

    def _send_click(self, point: tuple[int, int]) -> None:
        self.window.set_focus()
        self._send_mouse_message(WM_MOUSEMOVE, 0, point)
        time.sleep(0.02)
        self._send_mouse_message(WM_LBUTTONDOWN, MK_LBUTTON, point)
        time.sleep(0.03)
        self._send_mouse_message(WM_LBUTTONUP, 0, point)
        time.sleep(0.03)

    def _send_pick_and_place(self, src_point: tuple[int, int], dst_point: tuple[int, int]) -> None:
        self._send_click(src_point)
        time.sleep(0.04)
        self._send_click(dst_point)


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
