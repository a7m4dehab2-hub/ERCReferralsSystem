Vendored runtime libraries (no patient data is sent to a CDN):

- SheetJS Community Edition 0.20.3, official standalone distribution:
  https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
  License in SHEETJS-LICENSE.txt.
- Official @supabase/supabase-js 2.119.2, bundled locally with esbuild 0.28.2.
  Rebuild with `node build-vendor.mjs` after installing the locked dependencies.
  License text for the bundled dependencies is included in SUPABASE-LICENSES.txt.

The browser contacts only the configured Supabase project for sign-in and app data.
