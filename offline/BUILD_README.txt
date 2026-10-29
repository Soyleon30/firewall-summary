FirewallSummary offline Windows build package
=============================================

Purpose
-------
This package generates the final standalone Windows exe on a Windows target
computer. The target computer does not need an existing Python installation
and does not need network access.

Steps
-----
1. Extract the entire zip to a short local path, such as C:\FirewallBuild.
2. Double-click build.bat.
3. Wait for all four build stages to complete.
4. Find the final program at output\FirewallSummary.exe.

Included environment
--------------------
1. Official Windows x64 Python 3.12 installer.
2. Offline wheels for openpyxl, xlrd, PyInstaller, and Windows dependencies.
3. Application source and PyInstaller spec.
4. A packaged-exe self-test that verifies the embedded Python runtime and
   dependencies before copying the exe to output.

The generated FirewallSummary.exe is a one-file Windows program containing
Python, Tk, openpyxl, xlrd, and the application source.
