# SolidWorks Note Buddy

SolidWorks Note Buddy is a Python-based GUI application for annotating and managing notes on SolidWorks 3D models. This tool integrates SolidWorks with Tkinter and enables users to place markers, annotate frames, and export annotations as a LYNX file or PDF report.

## Features

- **3D Model Rotation and Capture**: Rotate models horizontally and vertically, capturing frames for annotation.
- **Marker Placement**: Click on the canvas to add markers, which can be annotated with custom notes.
- **Frame-by-Frame Navigation**: Use horizontal and vertical sliders to navigate through frames.
- **File Export**:
  - Export annotations to a PDF.
  - Save the project as a LYNX file, containing an animation of the annotated model.
- **Progress Tracking**: Track the rotation and capture progress with a progress bar.
- **Responsive UI**: Organized layout using Tkinter for easy navigation and control.

## Prerequisites

To run this script, you need the following libraries:
- `win32com.client`
- `Pillow`
- `reportlab`
- `tkinterdnd2`
- `json`
- `tkinter` (standard in Python)

## Installation

1. Clone or download this repository.
2. Install required packages:
   ```bash
   pip install pywin32 pillow reportlab tkinterdnd2
