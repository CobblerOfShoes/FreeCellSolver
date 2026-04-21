TARGET := find_cards.exe
SRC := read_memory.cpp
CXX := g++
CXXFLAGS :=
LDFLAGS := -lws2_32

.PHONY: all clean

all: $(TARGET)

find-cards:
	python AutoSolver.py

$(TARGET): $(SRC)
	$(CXX) $(CXXFLAGS) $(SRC) -o $(TARGET) $(LDFLAGS)

clean:
	del /Q $(TARGET)
