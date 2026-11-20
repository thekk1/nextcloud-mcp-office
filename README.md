# nextcloud-mcp-office

[nextcloud-mcp-server](https://github.com/cbcoutinho/nextcloud-mcp-server) 0.212.0 with additional tools:

- `bos_office.py`: Office tools (create, edit and convert documents, spreadsheets, presentations and PDF in
  Nextcloud), copying chat files into Nextcloud (`nc_office_save_chat_file`), `nc_files_check_list`, and
  `nc_fulltext_search` (content search through the Nextcloud full text index, app `fulltextsearch`), and
  `nc_activity_list` (who did what in a period, from the activity stream, app `activity`).
- `bos_start.py`: starts the server with the Office tools as an additional app `office`, a document processor
  for `.docx`, and an additional ownership check for Login Flow v2 grants.

The image is derived from the upstream image (pinned by digest, see `Dockerfile`).

```bash
docker build -t nextcloud-mcp-office .
```
