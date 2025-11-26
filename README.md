
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
  pip3 install pandas astropy astroquery==0.4.11
```
Note: Any version of astroquery 0.4.8 or newer should work, but tested on 0.4.11

Go to the project directory

```bash
  cd my-project
```

Clone the project

```bash
  git clone https://github.com/obaek/Astronomical-Distances-PI.git
```
Note: calculate_distances.py and object_type_map.py need to live in the same directory.

## Usage Guide

Step 1: Generate Data in PixInsight

Open your image in PixInsight.

Plate Solve the image (Script > Astrometry > ImageSolver). Crucial: The script will not work without a valid WCS solution.

Open the Annotation tool (Script > Astrometry > AnnotateImage).

Select your desired catalogs (Messier, NGC, Custom, etc.).

Check the "Write to file" box at the bottom of the dialog.

Click OK. Note the location of the generated objects.txt file.

Step 2: Run the script

```bash
  python3 calculate_distances.py [/path/to/file/]objects.txt
```

Optionally

```bash
  python3 calculate_distances.py
  Enter input filename: [/path/to/file/]objects.txt
```
If all goes well, Astronomical_Distances_TAP.xlsx will be written to the objects.txt directory.
