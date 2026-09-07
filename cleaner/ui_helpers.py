"""Small pure helpers shared by the Streamlit UI and its tests."""


def default_sheet_selection(sheet_names):
    """Return the CLI-compatible initial selection: only the first sheet."""
    return list(sheet_names[:1])
