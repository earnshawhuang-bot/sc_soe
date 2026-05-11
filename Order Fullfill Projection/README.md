# Data Input Folder Structure

This directory stores local monthly input files for the S&OE tracking workflow.

The folder structure is tracked in git so a fresh clone keeps the expected project layout. Real Excel input files are not tracked because they are monthly business data.

## Folders

| Folder | Purpose |
|--------|---------|
| `01-SC/` | Sales order baseline file, for example `Order tracking *.xlsx` |
| `02-Shipped/` | GI shipment file |
| `03-FG/` | Finished goods inventory snapshot |
| `04-PP/` | Scheduled and unscheduled production plan files |
| `06-Loading Plan/` | KS and IDN loading plan files |
| `07-Mapping/` | Local mapping/config workbooks such as customer mapping |

## Usage

After cloning the project on another computer, place the local Excel files into the matching folders and update `config.yaml` if the filenames or run month changed.

Only `.gitkeep` placeholders and this README are tracked. Do not commit raw monthly input files.
