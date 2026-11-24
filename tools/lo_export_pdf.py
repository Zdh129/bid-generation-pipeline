#!/usr/bin/env python3
"""LibreOffice 转 PDF：在用户配置中放入 Basic 宏，导出前更新目录等域。"""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

MODULE = r'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
Sub UpdateAndExport
    Dim sIn As String, sOut As String
    Dim oUrl As String, oOut As String
    Dim oDoc As Object
    Dim oIndexes As Object
    Dim i As Integer
    Dim props(0) As New com.sun.star.beans.PropertyValue
    Dim loadProps(1) As New com.sun.star.beans.PropertyValue
    sIn = Environ(&quot;LO_EXPORT_IN&quot;)
    sOut = Environ(&quot;LO_EXPORT_OUT&quot;)
    If sIn = &quot;&quot; Or sOut = &quot;&quot; Then
        Exit Sub
    End If
    oUrl = ConvertToURL(sIn)
    oOut = ConvertToURL(sOut)
    loadProps(0).Name = &quot;Hidden&quot;
    loadProps(0).Value = True
    loadProps(1).Name = &quot;UpdateDocMode&quot;
    loadProps(1).Value = 3
    oDoc = StarDesktop.loadComponentFromURL(oUrl, &quot;_blank&quot;, 0, loadProps())
    On Error Resume Next
    oDoc.Refresh
    oIndexes = oDoc.getDocumentIndexes()
    If Not IsNull(oIndexes) Then
        For i = 0 To oIndexes.getCount() - 1
            oIndexes.getByIndex(i).update()
        Next i
    End If
    oDoc.reformat
    On Error GoTo 0
    props(0).Name = &quot;FilterName&quot;
    props(0).Value = &quot;writer_pdf_Export&quot;
    oDoc.storeToURL(oOut, props())
    oDoc.close(True)
End Sub
</script:module>
'''

SCRIPT_XLB = '''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false">
 <library:element library:name="Module1"/>
</library:library>
'''

SCRIPT_XLC = '''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:libraries PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "libraries.dtd">
<library:libraries xmlns:library="http://openoffice.org/2000/library" xmlns:xlink="http://www.w3.org/1999/xlink">
 <library:library library:name="Standard" library:link="false"/>
</library:libraries>
'''

REGXCU = '''<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Writer/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>2</value></prop></item>
</oor:items>
'''


def prepare_profile(dest: Path) -> Path:
    if dest.exists():
        shutil.rmtree(dest, ignore_errors=True)
    basic = dest / "user" / "basic"
    std = basic / "Standard"
    std.mkdir(parents=True)
    (std / "Module1.xba").write_text(MODULE, encoding="utf-8")
    (std / "script.xlb").write_text(SCRIPT_XLB, encoding="utf-8")
    (basic / "script.xlc").write_text(SCRIPT_XLC, encoding="utf-8")
    (dest / "user" / "registrymodifications.xcu").write_text(REGXCU, encoding="utf-8")
    return dest


def convert_with_field_update(docx_path: Path, pdf_path: Path, timeout: int = 600) -> tuple[bool, str]:
    """返回 (宏导出是否成功, 日志). 失败时调用方应回退普通 soffice --convert-to pdf。"""
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    profile = Path(tempfile.mkdtemp(prefix="lo_prof_upd_"))
    try:
        prepare_profile(profile)
        env = os.environ.copy()
        env["SAL_USE_VCLPLUGIN"] = "svp"
        env["LO_EXPORT_IN"] = str(docx_path.resolve())
        env["LO_EXPORT_OUT"] = str(pdf_path.resolve())
        cmd = [
            "soffice", "--headless", "--nologo", "--nofirststartwizard", "--norestore",
            f"-env:UserInstallation=file://{profile}",
            "macro:///Standard.Module1.UpdateAndExport",
        ]
        proc = subprocess.run(cmd, check=False, timeout=timeout, env=env,
                              stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        log = (proc.stdout or b"").decode("utf-8", "replace")[-2000:] + "\n" + (proc.stderr or b"").decode("utf-8", "replace")[-2000:]
        ok = proc.returncode == 0 and pdf_path.is_file() and pdf_path.stat().st_size > 1000
        return ok, f"rc={proc.returncode} size={pdf_path.stat().st_size if pdf_path.is_file() else 0}\n{log}"
    finally:
        shutil.rmtree(profile, ignore_errors=True)
