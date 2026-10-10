# Project instructions

## Skills: ask before use

Before invoking any skill (including the HyperFrames plugin skills):

1. Tell the user the skill's name and why it is needed.
2. State its approximate on-invoke token cost (e.g. `talking-head-recut` ~22k, `hyperframes-core` ~4k).
3. Wait for the user's approval before loading it.

Reply to the user in Arabic.

## Installed but disabled plugins

- `hyperframes@hyperframes` (HeyGen, v0.8.145): installed but DISABLED to save ~4k tokens per session. It provides 21 video skills (HTML-to-video rendering, talking-head recut, captions, PR-to-video, slideshow, motion graphics, music-to-video, etc.).
- If the user asks for video work, remind them it exists and ask whether to enable it (`claude plugin enable hyperframes@hyperframes`, then a new session may be needed).
