import win32com.client as win32
import os

class Excel:
    def __init__(self, visible=False):
      try:
        self.xl = win32.gencache.EnsureDispatch("Excel.Application")
      except AttributeError:
        # Remove cache and try again.
        MODULE_LIST = [m.__name__ for m in sys.modules.values()]
        for module in MODULE_LIST:
          if re.match(r'win32com\.gen_py\..+', module):
            del sys.modules[module]
        shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
        self.xl = win32.gencache.EnsureDispatch("Excel.Application")

        if visible:
            self.xl.Visible = visible

    def __enter__(self):
        return self.xl

    def __exit__(self, exception_type, exception_value, traceback):
        if not self.xl.Visible:
            self.xl.Quit()

class Word:
    def __init__(self, visible=False):
      try:
        self.wd = win32.gencache.EnsureDispatch("Word.Application")
      except AttributeError:
        # Remove cache and try again.
        MODULE_LIST = [m.__name__ for m in sys.modules.values()]
        for module in MODULE_LIST:
          if re.match(r'win32com\.gen_py\..+', module):
            del sys.modules[module]
        shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
        self.wd = win32.gencache.EnsureDispatch("Word.Application")

        if visible:
            self.wd.Visible = visible

    def __enter__(self):
        return self.wd

    def __exit__(self, exception_type, exception_value, traceback):
        if not self.wd.Visible:
            self.wd.Quit()
