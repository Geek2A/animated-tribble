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

Remove with `.\Install-PurcheoPrintAgent.ps1 -Uninstall`.

Files live in `%ProgramData%\PurcheoPrintAgent` (`config.json`, `agent.log`). `config.json` options: `ServerUrl`, `Token`, `SumatraPath`, `LogPath`, `SetPaperSize` (default `true`), `PrintSettings` (SumatraPDF print settings, default `noscale`).

## Testing without the service

```powershell
.\PurcheoPrintAgent.ps1 -ConfigPath C:\ProgramData\PurcheoPrintAgent\config.json
```

## Tuning

- Paper size: the agent sets a custom page size on the printer from `media_width_mm` x `media_height_mm` via a print ticket. If the driver rejects it, a warning is logged and the driver default is used; set `SetPaperSize` to `false` and configure the size in the driver instead.
- Rotation and calibration are applied by the server; adjust them per agent and per device in Purcheo.
