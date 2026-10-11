# Render-only collection faces

This path extends the registration work in pull requests #1238 and #1242. It keeps default registration conservative.

The internal `registerEmbeddedFontFaces` environment accepts a trusted `renderOnlyCollectionPolicy` setting. Set `allowUnsignedDerivedFaces` to `true` and `requireSignedFonts` to `false` to request this path. Documents and embedded font metadata must not supply this setting. The host can pass the same explicit setting to `createDocxEditor({ renderOnlyCollectionPolicy, ... })`. The editor forwards it only to its local font registration. Omitting it retains the current default. React and Vue adapter prop forwarding is outside this change.

The helper reads bounded collection signature records and the selected face directory. It creates a separate SFNT for rendering and updates that file's checksums. The raw source bytes, source hash, source signature, and face index stay unchanged. The resulting SFNT is unsigned. Structural checks do not verify a digital signature.

Default environments still refuse signed nonzero faces. Environments that require signed fonts also refuse this path. Individual face-level signatures remain unsupported. Malformed directories, overlapping signature ranges, and invalid face indexes fail.

Public structural fixtures test bounds, checksums, source preservation, and registration disposal. They contain no private font files and do not establish real browser glyph parity. The host option does not change font admission or authentication. A browser must still accept the selected rendering file.
