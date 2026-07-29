"""Test bootstrap.

app.py builds a module-level `app = create_app()` on import (gunicorn's
entrypoint). create_app fails startup validation when the environment
carries neither APP_PASSWORD nor ALLOW_OPEN_ACCESS — correct in
production, but it would break merely importing the module under pytest.
Defaulting open access here affects only that import-time instance; every
test constructs its own app with explicit parameters.
"""

import os

os.environ.setdefault("ALLOW_OPEN_ACCESS", "1")
