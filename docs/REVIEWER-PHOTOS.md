# Hosting reviewer photos in this repository

Reviewer profiles are defined in `web/app.js` and copied to `docs/app.js` for GitHub Pages. Photos can be stored alongside the app instead of using an external image URL.

1. Obtain permission to use the photo and save an appropriately sized, web-friendly JPEG, PNG, or WebP. Use a descriptive, lowercase filename such as `gina-cecchetti.jpg`.
2. Create `web/assets/reviewers/` and `docs/assets/reviewers/` if they do not exist. Put the same photo in both directories, using the same filename.
3. In the reviewer's entry in `READER_PROFILES` in both `web/app.js` and `docs/app.js`, set the photo path relative to the app's `index.html`:

   ```js
   headshotUrl: "assets/reviewers/gina-cecchetti.jpg",
   ```

   For a new reviewer, also supply `fullName`, `firstName`, `aliases`, and `bio`. Add the names that may appear in the CSV to `aliases`; avoid aliases that belong to another reviewer.
4. Confirm the profile and photo files match in `web` and `docs`. Open the app, load a CSV containing the reviewer's name, and check the report preview and downloaded PDF for the correct photo and bio. Test both the full name and any first-name alias used in CSVs.
5. Commit and push both the code and photo files. GitHub Pages publishes the `docs` copy from the `main` branch; wait for publishing and refresh the app to verify the live result.

Existing profiles can continue to use external URLs until their photos are added to both directories and their `headshotUrl` values are updated. Adding this guide does not itself move any photos.

## Instructions for an AI assistant adding a photo

Inspect the actual image visually before choosing its framing. The app currently fills a square avatar using `object-fit: cover`, which centers the image by default. Check whether that crop shows the person's face clearly and leaves a natural amount of space around it.

If the default framing is unsuitable, add or use per-reviewer crop position and zoom settings in the app, then adjust them based on the image. Keep the original image unchanged. Apply matching changes to `web` and `docs`, and preserve the current default appearance for profiles without custom settings. The app does not currently have per-reviewer crop settings, so they must be implemented before custom framing can be applied.

Inspect the resulting avatar in the report preview and the downloaded PDF. Adjust the framing until the face is clear and the two outputs match. Do not assume the center of the source image is the best crop, and do not claim the result was visually verified unless you inspected it.
