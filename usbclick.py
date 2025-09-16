import usb.core
import usb.util
import time


import os
# os.environ['PYUSB_DEBUG'] = 'debug'

# usb.core.find()
# マウスデバイスを探す
dev = usb.core.find(idVendor=0x0458, idProduct=0x0003)  # ベンダーIDとプロダクトIDは実際のマウスに合わせて変更してください
# idVendor: 0x248A = TeLink Semiconductor (Shanghai) Co., Ltd.
# idProduct: 0x8367
if dev is None:
    raise ValueError("マウスデバイスが見つかりません")


dev.set_configuration()
cfg = dev.get_active_configuration()
intf = cfg[(0,0)]

ep_out = usb.util.find_descriptor(
    intf,
    custom_match=lambda e: usb.util.endpoint_direction(e.bEndpointAddress) == usb.util.ENDPOINT_OUT
)
if ep_out is None:
    raise ValueError("OUTエンドポイントが見つかりません")

def click():
    # 左クリック押下 (button=1, x=0, y=0, wheel=0)
    dev.write(ep_out.bEndpointAddress, [1, 0, 0, 0])
    time.sleep(0.1)
    # 離す
    dev.write(ep_out.bEndpointAddress, [0, 0, 0, 0])

click()
print("クリック送信完了")