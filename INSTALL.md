## File Structure

```
spire-presentation/
├── SKILL.md                      # Main skill definition (required)
├── references/                   # Detailed documentation (optional)
│   ├── 01-getting-started.md
│   ├── 02-basic-operations.md
│   ├── ... (15 reference files)
│   └── 15-best-practices.md
├── examples/                     # Code examples (optional)
│   ├── basic/
│   ├── charts/
│   ├── tables/
│   └── advanced/
├── evals/                        # Trigger evaluation (optional)
│   └── trigger_eval.json
└── INSTALL.md                    # This file
```

## Skill Metadata

```yaml
name: spire-presentation
description: This skill should be used when the user asks to "create a PowerPoint presentation", "edit a PPTX file", "convert PowerPoint to PDF", "add charts to slides", or mentions Spire.Presentation, PowerPoint automation, .NET presentation processing, or slide manipulation.
version: 0.1.0
```

## Usage

Once installed, the skill will automatically activate when you:
- Create or edit PowerPoint presentations
- Convert presentations to other formats
- Work with charts, tables, or SmartArt
- Add animations or multimedia
- Handle presentation security

## License

This skill references the Spire.Presentation .NET library. A separate license from E-iceblue is required for production use.

## Support

For Spire.Presentation documentation:
- https://www.e-iceblue.com/Introduce/presentation-for-net.html

For Claude Code skills information:
- https://docs.anthropic.com
