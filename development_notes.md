# Notes

* Specifying a hidden import to `pyinstaller`. From [here](https://stackoverflow.com/questions/47318119/no-module-named-pandas-libs-tslibs-timedeltas-in-pyinstaller)

  `pyinstaller --onefile --hidden-import pandas._libs.tslibs.timedeltas myScript.py`

