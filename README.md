# animated-tribble

PowerShell print agent for Purcheo. Polls the server for label and thermal print jobs and prints them on a local Windows printer.

## Install

1. In Purcheo, go to Settings > Printers > Print Agents, add an agent and copy the token (shown once).
2. Download the portable [SumatraPDF](https://www.sumatrapdfreader.org/) build (used for silent PDF printing).
3. In an elevated Windows PowerShell 5.1 prompt:

   ```powershell
   .\Install-PurcheoPrintAgent.ps1 -ServerUrl https://purcheo.example.com -SumatraPath C:\Tools\SumatraPDF.exe
   ```

   The token is prompted for if `-Token` is not given. Use `-Credential` to run the service as a user instead of LocalSystem (needed for printers installed per-user).
4. The agent reports its printers on first heartbeat; pick the printer in Purcheo.

Uninstall with `.\Uninstall-PurcheoPrintAgent.ps1` (stops and deletes the service and removes `%ProgramData%\PurcheoPrintAgent`). Add `-KeepConfig` to keep `config.json`; re-running `Install-PurcheoPrintAgent.ps1` with no parameters then reinstalls using the saved server URL and token. Re-running the installer over an existing install also works as an upgrade.

Files live in `%ProgramData%\PurcheoPrintAgent` (`config.json`, `agent.log`). `config.json` options: `ServerUrl`, `Token`, `SumatraPath`, `LogPath`, `SetPaperSize` (default `true`), `PrintSettings` (SumatraPDF print settings, default `noscale`), `PaperToleranceMm` (how far a driver paper size may differ from the label, default `1.5`), `SumatraTimeoutSeconds` (default `60`).

## Testing without the service

```powershell
.\PurcheoPrintAgent.ps1 -ConfigPath C:\ProgramData\PurcheoPrintAgent\config.json
```

## Tuning

- Paper size: the agent reads the paper sizes the printer's driver offers, picks the one closest to the job's `media_width_mm` x `media_height_mm` (within `PaperToleranceMm`, either orientation) and sets it on the printer. The first job logs the driver's full list. If nothing matches, or Windows rejects it, a warning with the details is logged and the driver default is used; set `SetPaperSize` to `false` and choose the size in the driver instead.
- Print check: after SumatraPDF hands the job over, the agent watches the Windows print queue and reports a failure if the job shows an error or stays queued for 30 s, instead of reporting success regardless.
- Rotation and calibration are applied by the server; adjust them per agent and per device in Purcheo.
