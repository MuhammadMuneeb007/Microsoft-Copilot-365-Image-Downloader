# 🖼️ Microsoft Copilot 365 Image Downloader 🤖

<div align="center">

![Microsoft Copilot](https://img.shields.io/badge/Microsoft-Copilot-blue?style=for-the-badge&logo=microsoft&logoColor=white)
![Python](https://img.shields.io/badge/Python-3.6+-yellow?style=for-the-badge&logo=python&logoColor=white)
![Windows](https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

</div>

<p align="center">
  <b>Automate image generation and download from Microsoft 365 Copilot with a single script!</b>
</p>

![Adaptable](https://github.com/user-attachments/assets/a2acc434-97cc-4682-a76a-777a91b67e51)

![image](https://github.com/user-attachments/assets/64b44f9b-0025-4e3e-b0a8-952dde71ad8f)


---

## ✨ Features

<table>
  <tr>
    <td>🔍</td>
    <td>Automatically detects and controls the Microsoft 365 Copilot window</td>
  </tr>
  <tr>
    <td>🗣️</td>
    <td>Creates new chats with targeted image generation prompts</td>
  </tr>
  <tr>
    <td>🎨</td>
    <td>Generates professional imagery for vocabulary terms</td>
  </tr>
  <tr>
    <td>💾</td>
    <td>Downloads, renames, and organizes images automatically</td>
  </tr>
  <tr>
    <td>🔄</td>
    <td>Processes entire vocabulary libraries in a single run</td>
  </tr>
</table>

## 🚀 Installation

```bash
# Clone this repository
git clone https://github.com/MuhammadMuneeb007/Microsoft-Copilot-Image-Downloader.git

# Navigate to the project directory
cd Microsoft-Copilot-Image-Downloader

# Install dependencies
pip install -r requirements.txt

# Run the script
python copilot_image_downloader.py
```

## 🛠️ Requirements

<details>
<summary>Click to expand requirements</summary>

- Windows operating system
- Microsoft 365 Copilot application
- Python 3.6+
- Internet connection
- Administrative privileges (recommended)
- Required Python packages:
  - pygetwindow
  - pyautogui
  - pywinauto
  - opencv-python
  - numpy
  - Pillow

</details>

## 🔄 How It Works

1. 🌐 **Initialization**
   - Script loads vocabulary dictionary
   - Creates download directory if needed

2. 🤖 **For Each Vocabulary Term**
   - Opens Microsoft 365 Copilot
   - Creates a new chat session
   - Sends customized prompt

3. ⏱️ **Processing**
   - Waits for image generation
   - Identifies UI elements using automation

4. 💾 **Download & Organization**
   - Clicks thumbnail to open image
   - Triggers download
   - Renames and moves file from Downloads folder
   - Organizes by vocabulary term

5. 🔁 **Repeat**
   - Process continues until all terms are completed

## 📚 Vocabulary Library

The script uses a comprehensive dictionary of vocabulary terms with carefully crafted image prompts:

<details>
<summary>Click to see the full vocabulary list</summary>

| Term | Image Prompt |
|------|--------------|
| **Adaptable** | A chameleon changing colors to blend with different environments |
| **Articulate** | Person speaking confidently to an audience with flowing words |
| **Collaborate** | Diverse team working together around a table |
| **Compassionate** | Person gently helping someone in need |
| **Creative** | Person with lightbulb surrounded by artistic tools |
| **Diligent** | Person working attentively on a task late into the night |
| **Efficient** | Well-organized factory with streamlined processes |
| **Eloquent** | Person giving a moving speech to a captivated audience |
| **Ethical** | Person making a difficult but honest decision |
| **Flexible** | Yoga practitioner in a challenging pose |
| **Genuine** | Person with warm smile and honest expression |
| **Gratitude** | Person expressing sincere thanks |
| **Innovative** | Bright lightbulb illuminating above a shiny cogwheel |
| **Integrity** | Person standing firmly on ethical principles |
| **Optimistic** | Person looking toward a bright, sunny future |
| **Persistent** | Person climbing a steep mountain despite obstacles |
| **Resourceful** | Person solving problems with limited materials |
| **Respectful** | Person listening attentively to someone else |
| **Responsible** | Person carefully tending to a garden |
| **Versatile** | Person juggling multiple tasks with ease |

</details>

## ⚙️ Customization

You can easily customize the vocabulary dictionary in the script:

```python
# Add your own vocabulary terms and prompts
vocab["Determined"] = "Show a person pushing against a boulder up a hill, sweating but resolute. Their face shows unwavering focus and determination."

# Remove existing terms
del vocab["Adaptable"]

# Modify existing prompts
vocab["Creative"] = "Show an explosion of colorful ideas coming from a person's mind, with various artistic elements floating around them."
```

## 🔍 Troubleshooting

![image](https://github.com/user-attachments/assets/7371fa80-e600-41d8-9e09-e74b857c4973)

<table>
  <tr>
    <th>Issue</th>
    <th>Solution</th>
  </tr>
  <tr>
    <td>Script can't find Copilot window</td>
    <td>
      <ul>
        <li>Open Copilot manually first</li>
        <li>Check if Copilot is installed</li>
        <li>Run script with admin privileges</li>
      </ul>
    </td>
  </tr>
  <tr>
    <td>UI elements not detected</td>
    <td>
      <ul>
        <li>Increase sleep times between actions</li>
        <li>Check console output for available UI elements</li>
        <li>Update to latest Copilot version</li>
      </ul>
    </td>
  </tr>
  <tr>
    <td>Download not working</td>
    <td>
      <ul>
        <li>Check Downloads folder permissions</li>
        <li>Ensure adequate disk space</li>
        <li>Try running with administrator rights</li>
      </ul>
    </td>
  </tr>
</table>

## 📊 Performance Tips

- 🖥️ Close unnecessary applications to reduce window detection issues
- ⚡ Running on a system with 8GB+ RAM improves performance
- 🔄 Restart your computer before running large batches
- 🕒 Adjust sleep times based on your system's performance and internet speed

## 📂 Project Structure

```
Microsoft-Copilot-Image-Downloader/
├── copilot_image_downloader.py  # Main script
├── requirements.txt             # Dependencies
├── README.md                    # Documentation
└── DownloadImages/              # Created automatically
    ├── Adaptable.png
    ├── Articulate.png
    └── ...
```

## 🤝 Contributing

Contributions are welcome! Feel free to:

- Report bugs and issues
- Suggest new features
- Submit pull requests
- Improve documentation

## 📄 License

This project is available under the MIT License. See the LICENSE file for details.

---

<div align="center">
  <p>
    <b>Developed by <a href="https://github.com/MuhammadMuneeb007">Muhammad Muneeb</a></b>
  </p>
  <p>
    <a href="https://github.com/MuhammadMuneeb007">GitHub</a>
  </p>
</div>
