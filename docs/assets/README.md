# Graphics

| File | What it is for |
| --- | --- |
| [hero.svg](hero.svg) | The animated explainer at the top of the README: a button pressed, the VBA converted, the copy updated. It loops every 12 seconds, and shows its last frame to people who turn animation off. |
| [hero.gif](hero.gif) | The same animation as a GIF, for places that don't play SVG: LinkedIn, X, Slack, slides. |
| [flyer.png](flyer.png) | A one-page explainer (2160 × 2700, portrait): what MacroPress does, how it works, why it's safe, how to try it. |
| [social-preview.png](social-preview.png) | The card shown when someone shares the repository's link. Upload it under the repository's **Settings → General → Social preview**. |

[make_graphics.py](make_graphics.py) draws `hero.svg`, `flyer.svg` and `social-preview.svg` from shared parts, so a
change to a colour or a line of copy lands in all three. Run it with `python docs/assets/make_graphics.py`, then
render the PNGs from the SVGs with any browser, as its docstring shows.
