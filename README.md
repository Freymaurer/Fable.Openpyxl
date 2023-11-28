# Fable.openpyxl
Fable bindings for the python xlsx reader writer openpyxl.


# Development

1. Create python virtual environment with `py -m venv .venv`
2. `dotnet tool restore`
3. `.\.venv\Scripts\python.exe -m pip install -r requirements.txt`, install local python dependencies

## Python Dependency Management

- Install new local dependencies with `.\.venv\Scripts\pip.exe install <PACKAGE_NAME>`
- Freeze local dependencies with `.\.venv\Scripts\python.exe -m pip freeze > requirements.txt` .
