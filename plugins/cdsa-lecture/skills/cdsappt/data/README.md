# Local corpus database

The public package does not include a lecture corpus. Build one from PowerPoint files you are allowed to use:

```powershell
powershell -ExecutionPolicy Bypass -File <plugin>\scripts\setup-corpus.ps1 -Source "C:\path\to\ppt-folder"
```

The script stores the database under `%LOCALAPPDATA%\cdsa-lecture` by default. Set `CDSA_LECTURE_DB` to that file before using corpus search and retrieval commands.
