"""Supported global eyedropper shortcuts and persisted-setting validation."""
# Global shortcuts must include Alt to avoid reserving common Ctrl commands.
MODIFIERS = {'Alt': 1, 'Ctrl+Alt': 3, 'Alt+Shift': 5, 'Ctrl+Alt+Shift': 7}
KEYS = {**{chr(v): v for v in range(65, 91)},
        **{str(v): 48 + v for v in range(10)},
        **{f'F{v}': 111 + v for v in range(1, 12)}}
DEFAULT_HOTKEY = {'enabled': True, 'modifiers': 'Alt', 'key': 'P'}


def clean_hotkey(value):
    if (not isinstance(value, dict)
            or type(value.get('enabled')) is not bool
            or not isinstance(value.get('modifiers'), str)
            or value['modifiers'] not in MODIFIERS
            or not isinstance(value.get('key'), str)
            or value['key'] not in KEYS):
        return dict(DEFAULT_HOTKEY)
    return {key: value[key] for key in DEFAULT_HOTKEY}


def hotkey_label(value):
    return value['modifiers'] + '+' + value['key']
