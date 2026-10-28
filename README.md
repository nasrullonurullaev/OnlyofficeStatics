# ONLYOFFICE Download Analytics

Collects lifetime download counts for tracked stable releases of
`ONLYOFFICE/DesktopEditors` and `ONLYOFFICE/DocumentServer` and publishes
`ONLYOFFICE-Download-Statistics.xlsx` to ONLYOFFICE DocSpace.
The first run creates a new report; later runs update that filename.
Existing files with other names are left untouched.

Two permanent visible sheets: **Desktop Editors** and **Document Server**.
The first run selects the two latest published stable releases per product.
Later runs refresh tracked releases and add releases published after the newest
tracked release. Older untracked releases are not imported; tracked history is retained.
Rows contain releases (newest first), categories, platforms, and artifact names.
Each run appends a date column on the right. New releases appear above older ones;
old counts move together with their artifacts. No new dated sheets are created.

History is matched by product, GitHub release ID, and asset ID, not row number or
filename. New artifacts show `—` for previous dates; missing artifacts retain their
history and show `—` for the current run. Zero is recorded as zero. Reuploaded files
with a new ID are separate rows. Category and platform totals sum recorded counts.
No separate summary, percentages, or hyperlinks.

Dates use `YYYY-MM-DD`; repeated runs add the time and a suffix if needed.
A hidden `_history` sheet stores identities and raw counts. Do not edit or delete it.
Start this format with a new document; old snapshot reports are not migrated.

## Setup

In **Settings → Secrets and variables → Actions**, add:

| Type | Name | Value |
| --- | --- | --- |
| Secret | `DOCSPACE_API_KEY` | API key with Files & Folders read/write access |
| Variable | `DOCSPACE_FOLDER_ID` | Destination folder ID, e.g. `17974` |
| Variable (optional) | `DOCSPACE_URL` | Defaults to `https://nct2.onlyoffice.com` |

The API key owner must have upload access to the folder. The workflow uses the
built-in GitHub token; no separate GitHub token is needed. The runner must be
able to reach DocSpace.

## Run

Open **Actions → ONLYOFFICE download statistics → Run workflow**.

Every manual and scheduled workflow run updates the document on nct2.

Scheduled runs update tracked releases every Monday at **04:17 UTC**
(09:17 Asia/Tashkent), from the default branch. GitHub may delay scheduled runs.
Reports are not committed or uploaded as Actions artifacts.

## Report rules

- `addon-*`: Add-ons; other `.zip` files: Updates; remaining packages: Installations.
- Installations include portable builds and Linux help packages.
- Unknown platforms, invalid counters, or duplicate asset IDs stop publication.
- Drafts, prereleases, and automatic source archives are excluded.
- Empty categories are omitted from file details. A collection error for either product stops publication.

## Upload errors

The current workbook is downloaded before adding the next date column.
Download or workbook-read errors stop publication to protect the existing history.
The script checks the file ID and version before and after uploading.
Duplicate names or a new file ID cause the run to fail. Resolve duplicates and
close the report in the editor before retrying. Uploads are not retried automatically.
DocSpace may retain document versions; the script does not delete them.

## Local use

Requires Python 3.10+.

```sh
python -m pip install -r requirements.in
python desktop_downloads.py
python -m unittest discover -s tests -v
```

Output: `reports/ONLYOFFICE-Download-Statistics.xlsx`. To publish, set the environment variables
listed above and add `--publish`. `.gitignore` excludes reports, Python caches,
and the local virtual environment. Local runs append a column to the existing output file;
publishing runs always use the workbook downloaded from DocSpace as their history.
