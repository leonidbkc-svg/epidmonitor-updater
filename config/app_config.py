import json
import os

DEFAULTS = {
    # РїРѕСЃС‚Р°РІСЊ СЃРІРѕРё РґРµС„РѕР»С‚РЅС‹Рµ СЃРµС‚РµРІС‹Рµ РїСѓС‚Рё (РјРѕР¶РЅРѕ РїРѕС‚РѕРј РјРµРЅСЏС‚СЊ РєРЅРѕРїРєР°РјРё)
    "swabs_archive_dir": r"\\server\share\archive\swabs",
    "swabs_journal_xlsx": r"\\server\share\Р›РђР‘Рђ Р’РЅСѓС‚СЂРµРЅРЅРёРµ РєРѕРЅС‚СЂРѕР»СЊ 2026.xlsx",
    "swabs_aliases_json": r"\\server\share\swabs_dep_aliases.json",
    "ai_agent_id": "",
    "ai_base_url": "",
    "ai_api_key": "",
    "tg_exam_base_url": "http://147.45.163.136",
    "tg_exam_report_api_key": "",
    "swabs_template_path": "",
    "swabs_report_dir": "",
    "departments": [],
    "webdav_url": "https://dav.epid-test.ru/",
    "webdav_user": "epiduser",
    "webdav_password": "1430194",
    "webdav_drive": "W:"
}

def _config_path() -> str:
    # 1) СЏРІРЅС‹Р№ РїСѓС‚СЊ С‡РµСЂРµР· env (РµСЃР»Рё Р·Р°РґР°РЅ)
    env_path = os.environ.get("EPID_MONITOR_CONFIG", "").strip()
    if env_path:
        return env_path

    # 2) РєРѕРЅС„РёРі СЂСЏРґРѕРј СЃ СЂР°Р±РѕС‡РµР№ РїР°РїРєРѕР№ (СѓРґРѕР±РЅРѕ РґР»СЏ EXE/Р·Р°РїСѓСЃРєР° РёР· РїР°РїРєРё РїСЂРѕРµРєС‚Р°)
    cwd_path = os.path.join(os.getcwd(), "config", "app_config.json")
    if os.path.exists(cwd_path):
        return cwd_path

    # 3) С…СЂР°РЅРёС‚СЃСЏ СЂСЏРґРѕРј СЃ СЌС‚РёРј РјРѕРґСѓР»РµРј
    return os.path.join(os.path.dirname(__file__), "app_config.json")


def get_config_path() -> str:
    return _config_path()

def load_config() -> dict:
    path = _config_path()
    if not os.path.exists(path):
        save_config(DEFAULTS)
        return dict(DEFAULTS)

    with open(path, "r", encoding="utf-8-sig") as f:
        data = json.load(f) or {}

    out = dict(DEFAULTS)
    out.update(data)
    return out

def save_config(cfg: dict) -> None:
    path = _config_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)





