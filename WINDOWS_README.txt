Firewall application summary tool
=================================

Runtime requirements
--------------------
1. Windows 10 or Windows 11, 64-bit.
2. Python, Microsoft Excel, and third-party packages are not required.
3. Run FirewallSummary.exe directly.

Usage
-----
1. Select the directory to scan. The default is D:\【维护部】.
2. Click "开始扫描".
3. Review the summary and duplicate markers.
4. Click "导出Excel" to save the result workbook.

Notes
-----
1. The executable includes Python, Tk, openpyxl, and xlrd.
2. A one-file executable can take several seconds to open the first time.
3. Windows SmartScreen may show "Unknown publisher" because the executable is
   not signed with a commercial code-signing certificate.
4. The application does not upload scanned directories or workbook contents.
