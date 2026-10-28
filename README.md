# ONLYOFFICE Download Analytics

Collects lifetime download counts for the latest stable releases of
`ONLYOFFICE/DesktopEditors` and `ONLYOFFICE/DocumentServer` and publishes
`ONLYOFFICE-Download-Statistics.xlsx` to ONLYOFFICE DocSpace.
The first run creates a new report; later runs update that filename.
Existing files with other names are left untouched.

One sheet shows Desktop Editors in columns A–E and Document Server in columns G–K.
Each product has its own release date and file details grouped into
**Installations**, **Updates**, and **Add-ons**. Each run replaces the previous
contents. Totals appear within each platform and category. No separate summary,
historical sheets, percentages, or hyperlinks.

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

- Leave `release_tag` empty for the latest stable release, or enter a tag such as `v10.0.0`.
- `server_release_tag` selects the Document Server release independently; empty means latest stable.
- Disable `publish` to build without uploading. DocSpace credentials are not required for this mode.

Scheduled runs use the latest stable release of each product every Monday at **04:17 UTC**
(09:17 Asia/Tashkent), from the default branch. GitHub may delay scheduled runs.
Reports are not committed or uploaded as Actions artifacts.

## Report rules

- `addon-*`: Add-ons; other `.zip` files: Updates; remaining packages: Installations.
- Installations include portable builds and Linux help packages.
- Unknown platforms, invalid counters, or duplicate asset IDs stop publication.
- Drafts, prereleases, and automatic source archives are excluded.
- Empty categories are omitted from file details. A collection error for either product stops publication.

## Upload errors

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
and the local virtual environment.
