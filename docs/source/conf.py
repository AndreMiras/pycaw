# Configuration file for the Sphinx documentation builder.
import os
import sys

# Add project root to path for autodoc
sys.path.insert(0, os.path.abspath('../..'))

# -- Project information -----------------------------------------------------
project = 'Pycaw'
copyright = '2025, Andre Miras'
author = 'Andre Miras'

# -- General configuration ---------------------------------------------------
extensions = [
    'sphinx.ext.autodoc',           # Auto-generate from docstrings
    'sphinx.ext.napoleon',          # Google/NumPy docstring support
    'sphinx.ext.viewcode',          # Add [source] links
    'sphinx.ext.intersphinx',       # Link to other docs (e.g., Python docs)
    'sphinx_autodoc_typehints',     # Better type hint rendering
]

# Autodoc settings
autodoc_default_options = {
    'members': True,
    'undoc-members': True,
    'show-inheritance': True,
}

# Napoleon settings for docstring parsing
napoleon_google_docstring = True
napoleon_numpy_docstring = True

# Intersphinx mapping (link to Python docs)
intersphinx_mapping = {
    'python': ('https://docs.python.org/3', None),
}

templates_path = ['_templates']
exclude_patterns = []

# -- Options for HTML output -------------------------------------------------
html_theme = 'sphinx_rtd_theme'
# html_static_path = ['_static']  # Commented out - directory is empty
html_title = f"{project} Documentation"

# Theme options
html_theme_options = {
    'navigation_depth': 4,
    'collapse_navigation': False,
}
