import os
import yaml


def load_config():
    config_path = os.path.join(os.path.dirname(__file__), "config", "config.yaml")
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def main():
    config = load_config()
    app_name = config.get("app_name", "app")
    env = config.get("env", "dev")
    print(f"{app_name} ({env})")


if __name__ == "__main__":
    main()
