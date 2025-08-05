# from pynput import mouse
# from PIL import Image
# import mss
# import time

# positions = []

# def on_click(x, y, button, pressed):
#     if pressed:
#         print(f"Clicked at: ({x}, {y})")
#         positions.append((x, y))
#         if len(positions) == 2:
#             return False  # Stop listening after two clicks

# def capture_selected_area():
#     print("Click TOP-LEFT corner of the region...")
#     print("Then scroll if needed, and click BOTTOM-RIGHT corner...")
#     print("Waiting for your clicks...")

#     # Start mouse listener
#     with mouse.Listener(on_click=on_click) as listener:
#         listener.join()

#     x1, y1 = positions[0]
#     x2, y2 = positions[1]

#     left = min(x1, x2)
#     top = min(y1, y2)
#     width = abs(x2 - x1)
#     height = abs(y2 - y1)

#     print(f"Capturing screen area: ({left}, {top}, {width}, {height})")
#     time.sleep(1)  # Give a moment before capture

#     with mss.mss() as sct:
#         monitor = {"top": top, "left": left, "width": width, "height": height}
#         img = sct.grab(monitor)
#         img_pil = Image.frombytes("RGB", img.size, img.rgb)
#         img_pil.save("selected_region.png")
#         print("✅ Screenshot saved as 'selected_region.png'")

# # Run the tool
# if __name__ == "__main__":
#     capture_selected_area()




from pynput import mouse
from PIL import Image
import mss
import time
import pyautogui  # For simulating scroll

positions = []

def on_click(x, y, button, pressed):
    if pressed:
        print(f"Clicked at: ({x}, {y})")
        positions.append((x, y))
        if len(positions) == 2:
            return False  # Stop listening after two clicks

def capture_scrolled_area():
    print("Click TOP-LEFT corner of the region...")
    print("Then scroll horizontally if needed, and click BOTTOM-RIGHT corner...")
    print("Waiting for your clicks...")

    # Start mouse listener
    with mouse.Listener(on_click=on_click) as listener:
        listener.join()

    x1, y1 = positions[0]
    x2, y2 = positions[1]

    left = min(x1, x2)
    top = min(y1, y2)
    width = abs(x2 - x1)
    height = abs(y2 - y1)

    print(f"Selected area: ({left}, {top}, {width}, {height})")
    time.sleep(1)  # Give a moment before capture

    # Initialize variables for scrolling
    screenshots = []
    scroll_increment = width  # Scroll by the width of the visible area
    total_width = width  # Initial width
    current_scroll = 0

    with mss.mss() as sct:
        while True:
            # Capture the current visible area
            monitor = {"top": top, "left": left, "width": width, "height": height}
            img = sct.grab(monitor)
            screenshots.append(Image.frombytes("RGB", img.size, img.rgb))

            # Simulate horizontal scroll (move right)
            pyautogui.scroll(-scroll_increment)  # Negative to scroll right
            current_scroll += scroll_increment
            time.sleep(0.5)  # Wait for scroll to settle

            # Check if we've scrolled enough (you can adjust this condition)
            # For simplicity, let's assume we scroll until we can't anymore
            # You may need to define a max scroll limit or detect scroll end
            if current_scroll >= width * 2:  # Arbitrary limit; adjust as needed
                break

    # Stitch images horizontally
    total_width = width * len(screenshots)
    final_image = Image.new("RGB", (total_width, height))
    for i, img in enumerate(screenshots):
        final_image.paste(img, (i * width, 0))

    # Save the final image
    final_image.save("scrolled_region.png")
    print("✅ Screenshot saved as 'scrolled_region.png'")

# Run the tool
if __name__ == "__main__":
    capture_scrolled_area()
    
    
    
    
    