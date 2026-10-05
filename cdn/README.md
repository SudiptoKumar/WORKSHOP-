# CDN Assets

This directory contains static assets intended to be served from the public GitHub repository through jsDelivr.

## URL pattern

```text
https://cdn.jsdelivr.net/gh/USER/REPOSITORY@VERSION/cdn/PATH
```

Examples:

```text
https://cdn.jsdelivr.net/gh/USER/REPOSITORY@v1.0.0/cdn/brand/finance-club.jpg
https://cdn.jsdelivr.net/gh/USER/REPOSITORY@v1.0.0/cdn/social/instagram.png
```

## Rules

- Keep production asset URLs pinned to an exact release/tag such as `@v1.0.0`.
- Keep asset paths lowercase and use kebab-case names.
- Keep permanent branding in `brand/`.
- Keep reusable social/channel icons in `social/`.
- Add new asset categories as top-level folders under `cdn/` instead of putting files directly in the root.
- Do not overwrite an existing production release to replace an asset. Publish a new version.
