# WORKSHOP CDN Assets

This directory contains static assets intended to be served directly from the repository through jsDelivr.

## jsDelivr URL pattern

```text
https://cdn.jsdelivr.net/gh/<GITHUB_USER>/<REPOSITORY>@<VERSION>/cdn/<PATH>
```

For production, use an immutable GitHub release/tag such as `v1.0.0` rather than `main`.

## Asset map

| Purpose | Path | Format |
|---|---|---|
| Finance Club brand image | `cdn/brand/finance-club.jpg` | JPG |
| Facebook icon | `cdn/social/facebook.png` | PNG |
| Instagram icon | `cdn/social/instagram.png` | PNG |
| LinkedIn icon | `cdn/social/linkedin.png` | PNG |
| X / Twitter icon | `cdn/social/twitter.png` | PNG |
| Certificate icon | `cdn/icons/certificate.png` | PNG |
| Registration icon | `cdn/icons/registration.png` | PNG |
| Calendar icon | `cdn/icons/calendar.png` | PNG |
| Business icon | `cdn/icons/business.png` | PNG |
| Google Meet icon | `cdn/integrations/meet.svg` | SVG |
| WhatsApp icon | `cdn/integrations/whatsapp.svg` | SVG |

## Naming rules

- Lowercase filenames.
- Kebab-case for multi-word names.
- Stable semantic names, not `final`, `new`, or numbered variants.
- Keep production assets in this tree; do not mix source code or templates into `cdn/`.
- When an existing production asset must be replaced, publish a new GitHub release version instead of changing the meaning of an existing immutable release.

## Example

```text
https://cdn.jsdelivr.net/gh/<GITHUB_USER>/<REPOSITORY>@v1.0.0/cdn/brand/finance-club.jpg
```
