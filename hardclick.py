import evdev
from evdev import UInput, ecodes as e

# マウスデバイスを見つける
devices = [evdev.InputDevice(path) for path in evdev.list_devices()]
mouse = None
for device in devices:
    if "mouse" in device.name.lower():
        mouse = device
        break

if mouse is None:
    print("マウスデバイスが見つかりません")
    exit()

# UInputを作成
ui = UInput.from_device(mouse)

# 左クリックを実行
ui.write(e.EV_KEY, e.BTN_LEFT, 1)  # マウスボタンを押下
ui.write(e.EV_KEY, e.BTN_LEFT, 0)  # マウスボタンを解放
ui.syn()

# UInputをクローズ
ui.close()