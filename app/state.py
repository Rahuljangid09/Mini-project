"""
Holds in-memory app state for the current single-user session.

Replaces the old bare module-level `analysis_data = {}` global + `global`
statement in main.py. Behavior is identical (still one shared dict, still
overwritten on every upload) — it's just no longer a stray global variable
scattered across the module, and it's easy to swap for real per-user
session storage later if the college guide asks for multi-user support.
"""


class AppState:
    def __init__(self):
        self.analysis_data: dict = {}

    def set_analysis_data(self, data: dict) -> None:
        self.analysis_data = data

    def get_analysis_data(self) -> dict:
        return self.analysis_data


# Single shared instance for this single-user app.
app_state = AppState()
