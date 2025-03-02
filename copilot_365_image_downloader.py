import pygetwindow as gw
import pyautogui
import time
import os
import cv2
import numpy as np
from PIL import ImageGrab, Image
# Add pywinauto imports
import pywinauto
from pywinauto import Desktop, Application

def open_microsoft_copilot():
    """Open Microsoft 365 Copilot application using Windows search"""
    print("Opening Microsoft 365 Copilot...")
    
    # First check if it's already open
    copilot_window = find_microsoft_copilot()
    if copilot_window:
        print("Microsoft 365 Copilot is already open")
        return copilot_window
    
    # If not open, try to launch it
    try:
        # Method: Using Windows search
        pyautogui.hotkey('win', 's')
        time.sleep(1)
        pyautogui.write('Microsoft 365 Copilot')
        time.sleep(1)
        pyautogui.press('enter')
        
        # Wait for the app to open
        print("Waiting for Copilot to open...")
        time.sleep(5)
        
        # Check if the window is now open
        return find_microsoft_copilot()
    
    except Exception as e:
        print(f"Error opening Microsoft 365 Copilot: {str(e)}")
        print("Please open Microsoft 365 Copilot manually and run the script again.")
        return None

def find_microsoft_copilot():
    """Find Microsoft 365 Copilot window using exact window name"""
    # Use the exact window name as shown in the window title
    window_name = "Microsoft 365 Copilot"
    
    # Get all windows
    windows = gw.getAllWindows()
    
    # Print all window titles for debugging
    print("All open windows:")
    for i, window in enumerate(windows):
        if window.title:
            print(f"{i+1}. {window.title}")
    
    # Find the exact match
    for window in windows:
        if window.title == window_name:
            print(f"Found exact match: {window.title}")
            return window
    
    # If exact match fails, try a partial match as fallback
    print("Exact match failed, trying partial match...")
    for window in windows:
        if window.title and window_name in window.title:
            print(f"Found partial match: {window.title}")
            return window
    
    print("No Microsoft 365 Copilot window found")
    return None

def focus_window(window):
    """Bring window to focus"""
    if window.isMinimized:
        window.restore()
    window.activate()
    time.sleep(1)  # Wait for window to come to foreground

def list_buttons_in_current_window():
    """List all buttons in the currently focused window"""
    print("Listing all buttons in the current window...")
    try:
        # Get the desktop object
        desktop = Desktop(backend="uia")
        
        # Get the foreground window
        main_window = desktop.window(active_only=True)
        
        if not main_window:
            print("Could not find active window")
            return []
            
        # Find all button controls
        buttons = main_window.descendants(control_type="Button")
        
        if not buttons:
            print("No buttons found in the current window")
            return []
            
        print(f"Found {len(buttons)} buttons:")
        button_list = []
        for i, button in enumerate(buttons):
            button_info = {
                'index': i+1,
                'text': button.window_text() if hasattr(button, 'window_text') else 'No text',
                'automation_id': button.automation_id() if hasattr(button, 'automation_id') else 'No ID',
                'rect': button.rectangle() if hasattr(button, 'rectangle') else 'No position'
            }
            button_list.append(button_info)
            print(f"{i+1}. Text: '{button_info['text']}', ID: {button_info['automation_id']}, Position: {button_info['rect']}")
        
        return button_list
    
    except Exception as e:
        print(f"Error listing buttons: {str(e)}")
        return []

def click_button_by_text(button_text):
    """Click a button with the specified text"""
    print(f"Looking for button with text: '{button_text}'")
    try:
        desktop = Desktop(backend="uia")
        main_window = desktop.window(active_only=True)
        
        if not main_window:
            print("Could not find active window")
            return False
            
        # Find all buttons
        buttons = main_window.descendants(control_type="Button")
        
        # Search for button with matching text
        for button in buttons:
            if button.window_text() == button_text:
                print(f"Found button: '{button_text}'")
                try:
                    button.click()
                    print(f"Successfully clicked button: '{button_text}'")
                    return True
                except Exception as click_error:
                    print(f"Error clicking button: {str(click_error)}")
                    return False
        
        print(f"No button found with text: '{button_text}'")
        return False
        
    except Exception as e:
        print(f"Error finding button: {str(e)}")
        return False

def click_button_by_text_if_it_contains(button_text):
    """Click a button with the specified text"""
    print(f"Looking for button with text: '{button_text}'")
    try:
        desktop = Desktop(backend="uia")
        main_window = desktop.window(active_only=True)
        
        if not main_window:
            print("Could not find active window")
            return False
            
        # Find all buttons
        buttons = main_window.descendants(control_type="Button")
        
        # Search for button with matching text
        for button in buttons:
            if button_text in  button.window_text() :
                print(f"Found button: '{button_text}'")
                try:
                    button.click()
                    print(f"Successfully clicked button: '{button_text}'")
                    return True
                except Exception as click_error:
                    print(f"Error clicking button: {str(click_error)}")
                    return False
        
        print(f"No button found with text: '{button_text}'")
        return False
        
    except Exception as e:
        print(f"Error finding button: {str(e)}")
        return False


def list_text_fields():
    """List all text input fields in the currently focused window"""
    print("Listing all text input fields in the current window...")
    try:
        desktop = Desktop(backend="uia")
        main_window = desktop.window(active_only=True)
        
        if not main_window:
            print("Could not find active window")
            return []
            
        # Find all edit controls
        edit_fields = main_window.descendants(control_type="Edit")
        
        if not edit_fields:
            print("No text input fields found in the current window")
            return []
            
        print(f"Found {len(edit_fields)} text input fields:")
        field_list = []
        for i, field in enumerate(edit_fields):
            field_info = {
                'index': i+1,
                'name': field.window_text() if hasattr(field, 'window_text') else 'No text',
                'automation_id': field.automation_id() if hasattr(field, 'automation_id') else 'No ID',
                'rect': field.rectangle() if hasattr(field, 'rectangle') else 'No position',
                'control': field  # Store the control for future interaction
            }
            field_list.append(field_info)
            print(f"{i+1}. Name: '{field_info['name']}', ID: {field_info['automation_id']}, Position: {field_info['rect']}")
        
        return field_list
    
    except Exception as e:
        print(f"Error listing text fields: {str(e)}")
        return []

def set_text_field_value(field_control, text):
    """Set text in a specific text field"""
    try:
        field_control.set_text(text)
        print(f"Successfully set text: '{text}'")
        return True
    except Exception as e:
        print(f"Error setting text: {str(e)}")
        return False

def click_and_set_text_field(field_control, text):
    """Click on a text field and then set its text using pyautogui"""
    try:
        # First click the field to focus it
        field_control.click_input()
        time.sleep(0.5)  # Small delay to ensure focus
        
        # Use pyautogui to write text
        pyautogui.write(text)
        time.sleep(0.5)
        # Press enter to send
        pyautogui.press('enter')
        print(f"Successfully clicked and sent text: '{text}'")
        return True
    except Exception as e:
        print(f"Error clicking and setting text: {str(e)}")
        return False

def get_downloaded_latest_image(key):
    # Get the downloads directory
    downloads_path = os.path.expanduser("~/Downloads")
    
    # Create DownloadImages directory if it doesn't exist
    download_images_dir = "DownloadImages"
    if not os.path.exists(download_images_dir):
        os.makedirs(download_images_dir)
    
    # Get the most recent file from downloads
    files = [os.path.join(downloads_path, f) for f in os.listdir(downloads_path) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
    if not files:
        print("No image files found in downloads")
        return
        
    latest_file = max(files, key=os.path.getctime)
    
    # Move the file to DownloadImages directory with the vocab key as name
     
    new_filename = os.path.join(download_images_dir, key+".png")
    
    try:
        os.rename(latest_file, new_filename)
        print(f"Image saved as {new_filename}")
    except Exception as e:
        print(f"Error saving image: {str(e)}")



vocab= {
  "Adaptable": "Show a chameleon changing its color to blend seamlessly with different environments. Alternatively, depict a person easily switching between different tasks or skills.",
  "Articulate": "Depict a person speaking confidently and clearly to a large audience. Their body language is engaging, and their words are represented as flowing smoothly and powerfully. Alternatively, show a well-written document with clear and concise sentences.",
  "Collaborate": "Show a diverse team of people working together around a table, brainstorming ideas and sharing information. Their expressions are engaged and collaborative. Use visual elements that suggest teamwork, cooperation, and shared success.",
  "Compassionate": "Show a person gently helping someone in need (e.g., offering food, providing comfort). The expression on the helper's face is kind and empathetic. Use colors and visual elements that convey warmth, caring, and understanding.",
  "Creative": "Show a person with a lightbulb above their head, surrounded by artistic tools and materials (e.g., paintbrushes, musical instruments, writing implements). The background is colorful and imaginative. Alternatively, show a collage of diverse artistic creations.",
  "Diligent": "Show a person working attentively and meticulously at a task, perhaps studying late into the night or carefully crafting something by hand. The setting is organized and focused, and the person's expression is determined.",
  "Efficient": "Show a well-organized factory or system where everything runs smoothly and productively. Use visual elements that suggest streamlined processes, optimal resource allocation, and minimal waste. Alternatively, show someone multitasking.",
  "Eloquent": "Show a person giving a moving and persuasive speech. Their words are powerful and inspiring, and the audience is captivated. Use visual elements that suggest grace, fluency, and artistry in communication.",
  "Ethical": "Show a person making a difficult but honest decision, even when it is unpopular or personally disadvantageous. The expression on their face is determined and principled. Use visual elements that suggest integrity, fairness, and moral courage.",
  "Flexible": "Show a yoga practitioner performing a challenging pose with ease and grace. Alternatively, depict a bamboo stalk bending in the wind without breaking. Use visual elements that suggest adaptability, resilience, and the ability to adjust to changing circumstances.",
  "Genuine": "Show a person with a warm smile and open, honest eyes. Their body language is relaxed and authentic. Avoid any visual elements that suggest deception or insincerity. Alternatively, someone who is not wearing any makeup or jewelery.",
  "Gratitude": "Show a person expressing thanks or appreciation to someone who has helped them. Their expression is sincere and heartfelt. Use visual elements that convey thankfulness, kindness, and positive emotion. Alternatively, show a gift.",
  "Innovative": "Show a lightbulb illuminating brightly above a gear or cogwheel. The cogwheel should be clean and shiny, turning. Smaller lightbulbs around the central one are dim or flickering, indicating older, less effective ideas. The overall image conveys a sense of creativity, progress, and forward-thinking.",
  "Integrity": "Show a person standing firmly on a foundation of honesty and ethical principles. They are surrounded by challenges and temptations, but they remain steadfast and true to their values. Use visual elements that suggest strength, trustworthiness, and moral uprightness.",
  "Optimistic": "Show a person looking towards a bright, sunny future with a hopeful smile. The sky is clear and blue, and the landscape is lush and inviting. Use visual elements that suggest positivity, hope, and belief in a better tomorrow. Alternatively, show someone who looks positive in the face of adversity.",
  "Persistent": "Show a person climbing a steep mountain, facing obstacles and setbacks along the way. They are determined to reach the summit, even when it is difficult and challenging. Use visual elements that suggest perseverance, resilience, and unwavering commitment.",
  "Resourceful": "Show a person using readily available materials and tools to solve a problem or create something new. They are inventive, adaptable, and able to make the most of limited resources. Use visual elements that suggest creativity, ingenuity, and problem-solving skills.",
  "Respectful": "Show a person listening attentively to someone else, valuing their opinions and perspectives. Their body language is open and receptive. Use visual elements that suggest courtesy, consideration, and valuing diversity.",
  "Responsible": "Show a person carefully tending to a garden or caring for an animal. They are reliable, dependable, and take their obligations seriously. Use visual elements that suggest commitment, accountability, and stewardship.",
  "Versatile": "Show a person juggling multiple tasks or skills with ease and proficiency. They are adaptable, flexible, and able to excel in a variety of roles. Use visual elements that suggest adaptability, breadth of knowledge, and skill proficiency."
}


for key,value in  vocab.items():
    
    # Step 1: Open Microsoft 365 Copilot
    print("Step 1: Opening Microsoft 365 Copilot...")
    copilot_window = open_microsoft_copilot()
        
    if not copilot_window:
        print("Failed to open or find Microsoft 365 Copilot window. Exiting.")
    else:
        # Step 2: Focus the window
        print("Step 2: Bringing window to focus...")
        focus_window(copilot_window)
        
        # Step 3: List all buttons in the window
        print("Step 3: Listing all buttons in the Copilot window...")
        buttons = list_buttons_in_current_window()
        
        if not buttons:
            print("No buttons were found or there was an error detecting buttons.")
        else:
            print(f"Successfully found {len(buttons)} buttons in the Microsoft 365 Copilot window.")
            # Example usage of clicking a button (replace 'Button Text' with actual button text)
            click_button_by_text("New chat")
        
        # # Step 3: List all text fields in the window
        print("Step 3: Listing all text input fields in the Copilot window...")
        text_fields = list_text_fields()
        
        if not text_fields:
            print("No text input fields were found or there was an error detecting fields.")
        else:
            print(f"Successfully found {len(text_fields)} text input fields in the Microsoft 365 Copilot window.")
            # Example usage:
            if text_fields:
                click_and_set_text_field(text_fields[0]['control'], value + "\n\ngenerate me a picture")
        
        time.sleep(60)
        buttons = click_button_by_text_if_it_contains("Thumbnail")
        time.sleep(10)
        buttons = click_button_by_text_if_it_contains("Download")
        time.sleep(10)
        buttons = click_button_by_text_if_it_contains("Close")
        time.sleep(10)
        buttons = click_button_by_text("")
        time.sleep(10)
        get_downloaded_latest_image(key)
