# docx-corruptor
This little Python script creates a .docx file that can't be opened because the document is corrupted and only contains random garbage anyway.
The created document retains it's file properties though, so from the outside the file appears to be intact.

# Configuration
Almost all (the most important) output file properties are configurable in the provided settings.cfg file.

# Execution
Just run docx-corruptor.py directly or via any console that supports Python. The configuration is loaded at runtime.

# Dependencies
Should run in any Python 3 environment. Only standard libraries are used: os, shutil, distutils, configparser, xml, random, datetime, re, zipfile.
