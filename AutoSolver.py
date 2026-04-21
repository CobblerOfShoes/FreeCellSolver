import argparse

from pywinauto import Application

class AutoSolver:
  def __init__(self, running: bool):
    if running:
      self.app: Application = Application(backend="win32").connect("FreeCell")
    else:
      self.app: Application = Application(backend="win32").start("Group_Project_T2/freecell.exe")
    self.window = self.app.window(title_re="FreeCell.*")

  def getControlIdentifiers(self):
    return self.window.print_control_identifiers()

  def startGame(self):
    ''' Starts a new game by sending the F2 key to the FreeCell window. '''
    #self.window.type_keys("{F2}")
    self.window.menu_select("Game->New Game")
    self.window = self.app.window(title_re="FreeCell.*")  # Update the window reference to the new game instance

  def quit(self):
    ''' Quits the game by sending the Alt+F4 keys to the FreeCell window. '''
    self.window.type_keys("%{F4}")
    confirmWindow = self.window.child_window(title="FreeCell", control_type="Window")
    if confirmWindow.exists(timeout=5):
      confirmWindow.child_window(title="Yes", control_type="Button").click()

def main():
  parser = argparse.ArgumentParser(
          prog='FreeCellSolver',
          description='Binds to a running instance of FreeCell and solves the game.',
  )
  parser.add_argument(
          '-r', '--running',
          type=bool,
          default=False,
          help='Binds to a running instance of FreeCell.',
  )

  args = parser.parse_args()

  solver = AutoSolver(args.running)

  solver.startGame()

  print(solver.getControlIdentifiers())

  input("Press Enter to quit the game...")
  solver.quit()

  print(solver.getControlIdentifiers())


if __name__ == "__main__":
  main()