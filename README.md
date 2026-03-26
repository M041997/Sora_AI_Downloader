## Sora Bulk Downloader (Drafts + Profile)

Small Windows desktop app to bulk-download **videos** and **thumbnails** from:

- `https://sora.chatgpt.com/drafts`
- `https://sora.chatgpt.com/profile`

### How it works

- You sign in inside the app (embedded browser).
- Click **Scan page** (it will auto-scroll a bit and also capture media URLs as you manually scroll).
- Click **Download** to save the queued videos/thumbnails to a folder you choose.

### Run (dev)

From your project folder (for example `C:\path\to\sora-bulk-downloader`):

```powershell
dotnet run --project .\SoraBulkDownloader.App\SoraBulkDownloader.App.csproj
```

### Build (Release)

```powershell
dotnet build .\SoraBulkDownloader.sln -c Release
```

Output:

- `SoraBulkDownloader.App\bin\Release\net8.0-windows\`

### Usage

1. Open the app.
2. Log in if prompted.
3. Click **Go: Drafts** or **Go: Profile**.
4. Click **Output folder…** and choose a directory.
5. Scroll the page (optional but recommended to load more items).
6. Click **Scan page**.
7. Click **Download**.

