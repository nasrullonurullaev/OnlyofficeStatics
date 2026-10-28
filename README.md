# ONLYOFFICE Download Analytics

Collects lifetime download counts for the latest and previous stable releases of
`ONLYOFFICE/DesktopEditors` and `ONLYOFFICE/DocumentServer` and publishes
`ONLYOFFICE-Download-Statistics.xlsx` to ONLYOFFICE DocSpace.
The first run creates a new report; later runs update that filename.
Existing files with other names are left untouched.

Each run appends a dated sheet to the same workbook, preserving earlier snapshots.
Desktop Editors occupies columns A–E and Document Server G–K. The latest releases
are at the top; the previous releases are below them. Every release has separate
file counts and platform/category totals: **Installations**, **Updates**, **Add-ons**.
There is no separate summary, percentages, or hyperlinks.

Sheet names use `YYYY-MM-DD`; repeated runs on the same date add the time and,
if needed, a numeric suffix. The newest sheet is selected when the file opens.
The previous release is the most recently published stable release before the
selected latest release. If none exists, only the available release is shown.

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

- Leave `release_tag` empty for the latest Desktop Editors release; a tag pins the top release.
- `server_release_tag` selects the top Document Server release independently.
- Each product's previous stable release is selected automatically.
- Disable `publish` to build without uploading. DocSpace credentials are not required for this mode.

Scheduled runs collect both releases of each product every Monday at **04:17 UTC**
(09:17 Asia/Tashkent), from the default branch. GitHub may delay scheduled runs.
Reports are not committed or uploaded as Actions artifacts.

## Report rules

- `addon-*`: Add-ons; other `.zip` files: Updates; remaining packages: Installations.
- Installations include portable builds and Linux help packages.
- Unknown platforms, invalid counters, or duplicate asset IDs stop publication.
- Drafts, prereleases, and automatic source archives are excluded.
- Empty categories are omitted from file details. A collection error for either product stops publication.

## Upload errors

The current workbook is downloaded before appending the next snapshot.
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
python -m unittest discover -s tests -p test_desktop_downloads.py -v
```

Output: `reports/ONLYOFFICE-Download-Statistics.xlsx`. To publish, set the environment variables
listed above and add `--publish`. `.gitignore` excludes reports, Python caches,
and the local virtual environment. Local runs append to the existing output file;
publishing runs always use the workbook downloaded from DocSpace as their history.
