# Security Policy

## Supported release

Security reports should target the latest published VFE Corel Production Tools 27 release.

## Macro safety

GMS/VBA projects can read and write files and modify open documents. Only load a GMS obtained from a trusted release. Verify the SHA-256 checksum and inspect source when appropriate.

Do not test new releases on the only copy of a production CDR file.

## Source-build trust setting

The temporary UI installer requires programmatic access to the active VBA project. Enable **Trust access to Visual Basic project** only during a controlled source build and disable it again afterward. End users installing the prebuilt GMS do not need this setting.

## Reporting

Report suspected vulnerabilities privately to VFE Flavius before public disclosure. Include the affected version, reproduction steps, impact, and a minimal sanitized sample when possible.

Do not include confidential customer documents in a public issue.
