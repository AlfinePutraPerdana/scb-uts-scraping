import json
import os
from core.utils import resource_path
PRESET_FILE = "presets.json"
WORKFLOW_PATH = "workflow.json"

def load_presets():
    path = resource_path(PRESET_FILE)
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_workflows():
    path = resource_path(WORKFLOW_PATH)
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)
    
def get_fields(team, screen, tab):
    presets = load_presets()
    try:
        return presets[team][screen][tab]
    except KeyError:
        return {}

def save_presets(presets):
    with open(PRESET_FILE, "w", encoding="utf-8") as f:
        json.dump(presets, f, indent=4)

def get_workflow(team, screen, tab):
    workflows = load_workflows()
    for wf in workflows:
        if wf["team"] == team and wf["screen"] == screen and wf["tab"] == tab:
            return wf["steps"]
    return []

def get_team_names():
    presets = load_presets()
    return list(presets.keys())


def get_available_teams():
    return list(load_presets().keys())

def get_screens_for_team(team):
    return list(load_presets().get(team, {}).keys())

def get_tabs_for_screen(team, screen):
    return list(load_presets().get(team, {}).get(screen, {}).keys())

def add_or_update_preset(team, screen, tab, fields):
    """
    Adds or updates the preset for a specific team/screen/tab.
    Fields is a dictionary of field_name -> {type, selector/...}
    """
    presets_data = load_presets()

    if team not in presets_data:
        presets_data[team] = {}

    if screen not in presets_data[team]:
        presets_data[team][screen] = {}

    presets_data[team][screen][tab] = fields

    save_presets(presets_data)


def delete_tab(team, screen, tab):
    data = load_presets()
    try:
        del data[team][screen][tab]
        if not data[team][screen]:
            del data[team][screen]
        if not data[team]:
            del data[team]
        save_presets(data)
    except KeyError:
        pass

def delete_screen(team, screen):
    data = load_presets()
    try:
        del data[team][screen]
        if not data[team]:
            del data[team]
        save_presets(data)
    except KeyError:
        pass

def delete_team(team):
    data = load_presets()
    try:
        del data[team]
        save_presets(data)
    except KeyError:
        pass

def get_all_presets():
    """
    Returns a flat list of (team, screen, tab) tuples for all presets.
    """
    presets = load_presets()
    flat = []
    for team, screens in presets.items():
        for screen, tabs in screens.items():
            for tab in tabs:
                flat.append((team, screen, tab))
    return flat
