# Media Placement Dashboard

This is a lightweight web dashboard that reads an Excel media placements file and visualizes:
- Mentions by media type
- Audience by media type
- Publicity by media type

Publicity is calculated per row using:
- Tier III -> `inches * 1000`
- Tier IV -> `inches * 1500`
- Tier V -> `inches * 3500`
- Tier VI -> `inches * 7500`

## Expected Excel columns
Column names are matched case-insensitively and can be close variants.

Required fields:
- `Media Type`
- `Audience`
- `Inches`
- `Tier`

## Run locally
Because browsers block local file module behavior for some environments, use a local server:

```bash
cd "/Users/marcelaerosa/Documents/New project"
python3 -m http.server 8080
```

Then open:
- <http://localhost:8080>

## Share with others
The app includes a **Copy Share Link** button.
- It stores the computed dashboard summary in the URL (`state` param).
- Anyone with the link can open and view the same metrics.

## Deploy for a public link
Fast options:
1. Netlify Drop: drag this folder into [Netlify Drop](https://app.netlify.com/drop)
2. Vercel: `vercel --prod`
3. GitHub Pages: push files and enable Pages for the repo
