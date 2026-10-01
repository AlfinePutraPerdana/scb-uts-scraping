from dataclasses import dataclass, field, asdict
from typing import List, Optional
import json
import os
from core.utils import resource_path, bundled_resource_path

@dataclass
class WorkflowStep:
    action: str  # e.g., "click", "input", "wait", "extract"
    selector_key: Optional[str] = None  #Optional e.g., "EUC/Relationship Individual/Main/RELID"
    params: dict = field(default_factory=dict)
    auto_generated: bool = False  # <-- NEW FIELD
    def to_dict(self):
        return asdict(self)

    @staticmethod
    def from_dict(data):
        return WorkflowStep(
            action=data["action"],
            selector_key=data.get("selector_key"),  # ✅ safe access
            params=data.get("params", {})
        )

@dataclass
class Workflow:
    name: str
    team: str
    screen: str
    tab: str
    steps: List[WorkflowStep] = field(default_factory=list)

    def to_dict(self):
        return {
            "name": self.name,
            "team": self.team,
            "screen": self.screen,
            "tab": self.tab,
            "steps": [step.to_dict() for step in self.steps]
        }

    @classmethod
    def from_dict(cls, data):
        return cls(
            name=data.get("name", "Unnamed Workflow"),
            team=data.get("team", ""),
            screen=data.get("screen", ""),
            tab=data.get("tab", ""),
            steps=[WorkflowStep.from_dict(step) for step in data.get("steps", [])]
        )

    def save_to_file(self, path):
        # always save into exe folder, never bundled temp
        full_path = resource_path(path)
        with open(full_path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=2)


    @classmethod
    def load_from_file(cls, path):
        with open(resource_path(path), "r", encoding="utf-8") as f:
            data = json.load(f)
        return cls.from_dict(data)

    @classmethod
    def load_all_workflows(cls, path) -> List["Workflow"]:
        # always prefer external (exe folder)
        full_path = resource_path(path)

        # fallback: bundled copy (if user deleted external accidentally)
        if not os.path.exists(full_path):
            full_path = bundled_resource_path(path)

        print("DEBUG: loading workflows from", full_path)

        with open(full_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        if not isinstance(data, list):
            raise ValueError("workflow.json should contain a list of workflows.")

        return [cls.from_dict(workflow_dict) for workflow_dict in data]

