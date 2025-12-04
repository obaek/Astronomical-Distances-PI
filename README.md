
# Astronomical Distances

This Python script generates an Excel spreadsheet of distances to astronomical objects obtained from Pixinsight's AnnotateImage script.

The script makes Simbad TAP queries and returns the following information:

1. Object Name
2. Object Type
3. RA
4. Dec 
5. Magnitude
6. Redshift
7. Distance (Parsecs)
8. Distance (Million Light Years)
9. Method

## Installation

Install required Python modules:

```bash
  pip3 install pandas tkinterdnd openpyxl astropy astroquery==0.4.11
```
Note: Any version of astroquery 0.4.8 or newer should work, but tested on 0.4.11

Download the files (or clone the project) and place all files in the same directory.

## Usage Guide

Step 1: Generate Data in PixInsight

Open your image in PixInsight.

Plate Solve the image (Script > Astrometry > ImageSolver). Crucial: The script will not work without a valid WCS solution.

Open the Annotation tool (Script > Astrometry > AnnotateImage).

Select your desired catalogs (Messier, NGC, Custom, etc.).

Check the "Write to file" box at the bottom of the dialog.

Click OK. Note the location of the generated objects.txt file.

Step 2: Usage (command line)

```bash
  python3 calculate_distances.py [-h] [--output OUTPUT] [input_file]
```
positional arguments:
  input_file       Path to the input text file containing object names

options:
  -h, --help       show this help message and exit
  --output OUTPUT  Path to the output Excel file

Graphical User Interface

```bash
  python3 calculate_distances_gui
```

