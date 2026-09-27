from pathlib import PureWindowsPath

import recent_office_finder as finder_module


def test_navigation_matches_windows_short_and_long_paths(monkeypatch):
    """A folder entered via its long path must match indexed 8.3-path children."""
    short_parent = PureWindowsPath(r"C:\Users\USER~1\Documents\workspace")
    long_parent = PureWindowsPath(r"C:\Users\sampleuser\Documents\workspace")
    child = finder_module.FinderItem(short_parent / "fixtures", True, 0)

    monkeypatch.setattr(
        finder_module.os.path,
        "realpath",
        lambda path: str(path).replace("USER~1", "sampleuser"),
    )

    app = finder_module.RecentOfficeFinder.__new__(finder_module.RecentOfficeFinder)
    app.all_items = [child]
    app.current_directory = long_parent

    assert app._get_navigation_items() == [child]
