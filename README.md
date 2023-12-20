# P3NE - PowerPoint Presenter Note Extractor

## Introduction

PowerPoint Presenter Note Extractor (P3NE) is a simple, user-friendly application that allows users to extract presenter notes from PowerPoint presentations (.pptx files) and save them as text files. This tool is designed to be intuitive and efficient, making it easy for anyone to quickly retrieve notes from their presentations.
Features
- Extract notes from PowerPoint slides.
- Save notes as a text file.
- Simple and intuitive graphical user interface.
- Standalone application for both Windows and macOS.

## Installation
#### Windows

- Download the latest release for Windows from the Releases section.
- Unzip the downloaded file.
- Run P3NE.exe to start the application.

#### macOS
- Download the latest release for macOS from the Releases section.
- Open the downloaded file (.dmg or .zip).
- Drag the P3NE app to your Applications folder.
- Double-click the application to run it. (If you encounter any security warnings since this is an unsigned application, refer to the FAQ section below.)

## Usage

1) Launch the P3NE application.
2) Click on the 'Browse' button to select a PowerPoint file.
3) After the notes are extracted, choose a location to save the text file.
4) The application will close once the file is saved.

## Building from Source

If you prefer to build the application from the source code:

- Clone the repository:

```
git clone https://github.com/paulbuchan/P3NE.git
```

- Navigate to the cloned directory and install the required dependencies:

```python 
pip install -r requirements.txt
```

- Run pyinstaller to build the executable:
  - For Windows: `pyinstaller --onefile P3NE.py`
  - For macOS: `pyinstaller --onefile P3NE.py`

## FAQ

_I'm getting a security warning on macOS. What should I do?_
  
macOS may block unsigned applications downloaded from the internet. To override this, locate the application in Finder, control-click the app icon, then select 'Open' from the shortcut menu.

## License
This project is licensed under the [MIT License](https://github.com/paulbuchan/P3NE/blob/main/License.md).


