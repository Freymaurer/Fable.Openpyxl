import openpyxl as openpyxl_1
from typing import Any
from fable_modules.fable_library.string_ import (to_console, printf)

openpyxl: Any = openpyxl_1

to_console(printf("%A"))(openpyxl)

