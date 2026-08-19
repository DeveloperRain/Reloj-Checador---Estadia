"""Lectura de asistencias mediante el SDK COM oficial de ZKTeco en Windows."""
import base64
import json
import os
import subprocess
from datetime import datetime
from types import SimpleNamespace


class OfficialSdkError(Exception):
    pass


def get_attendance(ip: str, port: int, timeout: int = 300):
    """Obtiene eventos con zkemkeeper, el mismo SDK usado por ZKTime.Net."""
    if os.name != "nt":
        raise OfficialSdkError("El SDK oficial de ZKTeco solo está disponible en Windows")

    powershell = os.path.join(
        os.environ.get("WINDIR", r"C:\\Windows"),
        "SysWOW64", "WindowsPowerShell", "v1.0", "powershell.exe",
    )
    if not os.path.isfile(powershell):
        raise OfficialSdkError("No se encontró PowerShell de 32 bits para zkemkeeper")

    script = r'''
$ErrorActionPreference = 'Stop'
$zk = New-Object -ComObject 'zkemkeeper.ZKEM.1'
if (-not $zk.Connect_Net($env:TIMECORE_ZK_IP, [int]$env:TIMECORE_ZK_PORT)) { throw 'El SDK oficial no pudo conectar al reloj' }
try {
  if (-not $zk.ReadAllGLogData(1)) { throw 'El reloj no inició la lectura de eventos' }
  $rows = [System.Collections.Generic.List[object]]::new()
  $pin=''; $verify=0; $inout=0; $year=0; $month=0; $day=0; $hour=0; $minute=0; $second=0; $workcode=0
  while ($zk.SSR_GetGeneralLogData(1,[ref]$pin,[ref]$verify,[ref]$inout,[ref]$year,[ref]$month,[ref]$day,[ref]$hour,[ref]$minute,[ref]$second,[ref]$workcode)) {
    $rows.Add([pscustomobject]@{user_id="$pin";status=$inout;punch=$verify;timestamp=('{0:D4}-{1:D2}-{2:D2}T{3:D2}:{4:D2}:{5:D2}' -f $year,$month,$day,$hour,$minute,$second)})
  }
  [Console]::Out.WriteLine(('TIMECORE_JSON=' + ($rows | ConvertTo-Json -Compress)))
} finally { $zk.Disconnect() | Out-Null }
'''
    encoded = base64.b64encode(script.encode("utf-16le")).decode("ascii")
    env = os.environ.copy()
    env.update({"TIMECORE_ZK_IP": str(ip), "TIMECORE_ZK_PORT": str(int(port))})
    result = subprocess.run(
        [powershell, "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded],
        capture_output=True, text=True, timeout=max(30, int(timeout)), env=env,
    )
    line = next((line for line in result.stdout.splitlines() if line.startswith("TIMECORE_JSON=")), None)
    if result.returncode != 0 or line is None:
        raise OfficialSdkError((result.stderr or result.stdout or "Error del SDK oficial").strip())

    data = json.loads(line.removeprefix("TIMECORE_JSON="))
    if isinstance(data, dict):
        data = [data]
    return [
        SimpleNamespace(
            user_id=str(row["user_id"]), uid=None,
            timestamp=datetime.fromisoformat(row["timestamp"]),
            status=int(row["status"]), punch=int(row["punch"]),
        )
        for row in data
    ]
