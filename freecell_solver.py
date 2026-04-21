#!/usr/bin/env python3
'''
Run this, then C++ reader with freecell game running in background

'''

import heapq
import socket
import struct
import collections

CARD_PORT = 5000
FRAME_SIZE = 772

rank_symbols = ['A','2','3','4','5','6','7','8','9','0','J','Q','K']
suit_symbols = ['S','D','H','C']

def decode_card(v):
    if v == -1 or v == 0xffffffff:
        return None
    if v < 0 or v > 51:
        return None
    rank = v // 4
    suit = v % 4
    return rank_symbols[rank] + suit_symbols[suit]

def decode_foundations(raw_foundations):
    ranks = [0, 0, 0, 0]
    for slot_idx, v in enumerate(raw_foundations):
        if v == -1 or v == 0xffffffff:
            continue
        if 0 <= v <= 51:
            suit_idx = v % 4
            rank = (v // 4) + 1
            if ranks[suit_idx] and ranks[suit_idx] != rank:
                print(
                    f"Warning: conflicting foundation values for suit {suit_idx}: "
                    f"{ranks[suit_idx]} vs {rank} from slot {slot_idx}"
                )
            ranks[suit_idx] = max(ranks[suit_idx], rank)
        else:
            print(f"Warning: invalid foundation value {v} in slot {slot_idx}")
    return tuple(ranks)

def read_exact(sock, size):
    buf = b''
    while len(buf) < size:
        chunk = sock.recv(size - len(buf))
        if not chunk:
            raise RuntimeError("Socket closed")
        buf += chunk
    return buf

def read_board_from_socket():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as server:
        server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        server.bind(("127.0.0.1", CARD_PORT))
        server.listen(1)

        print("Waiting for C++ reader...")

        sock, addr = server.accept()
        print("Connected from", addr)

        with sock:
            data = read_exact(sock, FRAME_SIZE)

    ints = struct.unpack("<193i", data)

    cards = ints[:189]
    raw_foundations = ints[189:193]
    foundations = decode_foundations(raw_foundations)
    freecells = []

    cols = [[] for _ in range(8)]

    for i in range(4):
        card = decode_card(cards[i])
        if card:
            freecells.append(card)

    for col in range(1, 9):
        for offset in range(21):
            v = cards[col*21 + offset]
            card = decode_card(v)
            if not card:
                break
            cols[col-1].append(card)

    all_cards = [c for col in cols for c in col] + list(freecells)
    duplicates = [c for c, n in collections.Counter(all_cards).items() if n > 1]
    if duplicates:
        print("Warning: duplicate cards detected in payload:", duplicates)
    if len(all_cards) + sum(foundations) != 52:
        print(
            "Warning: card count mismatch. columns+freecells+foundations =",
            len(all_cards) + sum(foundations),
        )

    return cols, freecells, foundations


class FreeCellSolver:
    def __init__(self, initial_columns, initial_freecells, foundations):
        self.initial_state = (
            tuple(tuple(c) for c in initial_columns),
            frozenset(initial_freecells),
            tuple(foundations)
        )

        self.suits_idx = {'S':0,'D':1,'H':2,'C':3}
        self.idx_suits = {0:'S',1:'D',2:'H',3:'C'}

        self.rank_map = {str(i): i for i in range(2,10)}
        self.rank_map.update({'A':1,'0':10,'J':11,'Q':12,'K':13})

    def heuristic(self, state):
        return (52 - sum(state[2])) + (len(state[1]) * 0.1)

    def is_red(self, suit):
        return suit in ('H','D')

    def can_stack(self, card, target):
        c_rank, c_suit = self.rank_map[card[:-1]], card[-1]
        t_rank, t_suit = self.rank_map[target[:-1]], target[-1]
        return (self.is_red(c_suit) != self.is_red(t_suit)) and (c_rank == t_rank - 1)

    def get_moves(self, state):
        cols, free, found = state
        cols = [list(c) for c in cols]
        free = set(free)
        found = list(found)
        possible = []

        for i, col in enumerate(cols):
            if col:
                card = col[-1]
                rank, s_idx = self.rank_map[card[:-1]], self.suits_idx[card[-1]]
                if rank == found[s_idx] + 1:
                    new_cols = [list(c) for c in cols]
                    new_cols[i].pop()
                    new_found = list(found)
                    new_found[s_idx] = rank
                    return [(tuple(tuple(c) for c in new_cols), frozenset(free), tuple(new_found), f"Move {card} from Col {i+1} to Foundation")]

        for card in free:
            rank, s_idx = self.rank_map[card[:-1]], self.suits_idx[card[-1]]
            if rank == found[s_idx] + 1:
                new_free = set(free)
                new_free.remove(card)
                new_found = list(found)
                new_found[s_idx] = rank
                return [(tuple(tuple(c) for c in cols), frozenset(new_free), tuple(new_found), f"Move {card} from FreeCell to Foundation")]

        for card in free:
            for i, col in enumerate(cols):
                if not col or self.can_stack(card, col[-1]):
                    new_cols = [list(c) for c in cols]
                    new_cols[i].append(card)
                    new_free = set(free)
                    new_free.remove(card)
                    dest = f"Col {i+1}" if col else f"empty Col {i+1}"
                    possible.append((tuple(tuple(c) for c in new_cols), frozenset(new_free), tuple(found), f"Move {card} from FreeCell to {dest}"))

        for i, col_from in enumerate(cols):
            if not col_from: continue
            card = col_from[-1]
            for j, col_to in enumerate(cols):
                if i == j: continue
                if not col_to or self.can_stack(card, col_to[-1]):
                    new_cols = [list(c) for c in cols]
                    new_cols[j].append(new_cols[i].pop())
                    dest = f"Col {j+1}" if col_to else f"empty Col {j+1}"
                    possible.append((tuple(tuple(c) for c in new_cols), frozenset(free), tuple(found), f"Move {card} from Col {i+1} to {dest}"))

        if len(free) < 4:
            for i, col in enumerate(cols):
                if col:
                    new_cols = [list(c) for c in cols]
                    card = new_cols[i].pop()
                    new_free = set(free)
                    new_free.add(card)
                    possible.append((tuple(tuple(c) for c in new_cols), frozenset(new_free), tuple(found), f"Move {card} from Col {i+1} to FreeCell"))

        return possible

    def solve(self):
        weight = 15.0
        start_node = (weight*self.heuristic(self.initial_state),0,self.initial_state,[])
        pq=[start_node]
        visited={self.initial_state:0}

        count=0
        while pq:
            _,steps,state,path=heapq.heappop(pq)
            count+=1

            if count%10000==0:
                print(f"Searching... States:{count} | Foundation:{sum(state[2])}/52",end="\r")

            if state[2]==(13,13,13,13):
                print(f"\nSolution found in {count} iterations")
                return path

            for next_state,next_free,next_found,move_text in self.get_moves(state):
                new_state=(next_state,next_free,next_found)
                new_steps=steps+1

                if new_state not in visited or new_steps<visited[new_state]:
                    visited[new_state]=new_steps
                    priority=new_steps+(weight*self.heuristic(new_state))
                    heapq.heappush(pq,(priority,new_steps,new_state,path+[move_text]))

        return None


if __name__ == "__main__":

    cols, freecells, foundations = read_board_from_socket()

    solver = FreeCellSolver(cols, freecells, foundations)

    print("Solving...")

    solution_path = solver.solve()

    if solution_path:
        print("\n--- STEP-BY-STEP SOLUTION ---")
        for i,move in enumerate(solution_path,1):
            print(f"{i}. {move}")
        print("\nGame Over: All cards in Foundation!")
        print("\nsaving solution to solution.txt...")
        with open("solution.txt", "w") as f:
            for i,move in enumerate(solution_path,1):
                f.write(f"{i}. {move}\n")
    else:
        print("\nNo solution found.")
