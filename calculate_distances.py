import pandas as pd
from astroquery.simbad import Simbad
from astropy import units as u
from astropy.coordinates import SkyCoord
import numpy as np
import os
import argparse
import time

# Constants
c_kms = 299792.458  # Speed of light in km/s
H0 = 70.0           # Hubble constant in km/s/Mpc
MPC_TO_LY = 3261564 # Conversion factor from Mpc to Light Years
LY_TO_MLY = 1e-6    # Conversion factor from Light Years to Million Light Years
KPC_TO_MPC = 1e-3   # Conversion factor from kpc to Mpc

# Object Type Mapping
try:
    from object_type_map import OBJECT_TYPE_MAP
except ImportError:
    print("Warning: object_type_map.py not found. Using fallback map.")
    OBJECT_TYPE_MAP = {}

def get_object_data_tap(object_name):
    """
    Queries Simbad via TAP for the object's data.
    Returns (redshift, distance_mly, distance_pc, method, ra, dec, magnitude, object_type)
    """
    try:
        s = Simbad()
        
        # Handle PK object formatting: PKnnn+nn.n -> PKnnn+nn n
        query_name = object_name
        if object_name.startswith('PK'):
            if '.' in query_name:
                parts = query_name.rsplit('.', 1)
                if len(parts) == 2:
                    query_name = f"{parts[0]} {parts[1]}"

        # Construct ADQL Query
        # We join basic, ident, allfluxes, and mesDistance
        # We fetch ALL distance measurements to filter for parallax or most recent
        query = f"""
        SELECT
            basic.main_id, basic.ra, basic.dec, basic.otype,
            basic.rvz_radvel, basic.rvz_redshift,
            flux.V,
            dist.dist, dist.unit, dist.method, dist.bibcode
        FROM basic
        JOIN ident ON ident.oidref = basic.oid
        LEFT JOIN allfluxes AS flux ON flux.oidref = basic.oid
        LEFT JOIN mesDistance AS dist ON dist.oidref = basic.oid
        WHERE ident.id = '{query_name}'
        """
        
        result = s.query_tap(query)
        
        if result is None or len(result) == 0:
            # Fallback: Try resolving name via standard query to get Main ID, then query TAP
            try:
                # Reset fields to minimal to avoid deprecation warnings
                s.reset_votable_fields()
                s.add_votable_fields('oid')
                r_resolve = s.query_object(query_name)
                if r_resolve is not None and len(r_resolve) > 0 and 'oid' in r_resolve.colnames:
                    oid = r_resolve['oid'][0]
                    # Query by OID
                    query_oid = f"""
                    SELECT
                        basic.main_id, basic.ra, basic.dec, basic.otype,
                        basic.rvz_radvel, basic.rvz_redshift,
                        flux.V,
                        dist.dist, dist.unit, dist.method, dist.bibcode
                    FROM basic
                    LEFT JOIN allfluxes AS flux ON flux.oidref = basic.oid
                    LEFT JOIN mesDistance AS dist ON dist.oidref = basic.oid
                    WHERE basic.oid = {oid}
                    """
                    result = s.query_tap(query_oid)
            except Exception as e_resolve:
                print(f"Resolution fallback failed for {object_name}: {e_resolve}")

        if result is None or len(result) == 0:
            print(f"Object {object_name} not found in Simbad TAP.")
            return None, None, None, "No Data", None, None, None, None
            
        # Use the first row for basic info (RA, Dec, Mag, Redshift)
        row = result[0]
        
        # Get Coordinates
        ra = None
        dec = None
        if 'ra' in result.colnames and 'dec' in result.colnames:
            ra_deg = row['ra']
            dec_deg = row['dec']
            if not np.ma.is_masked(ra_deg) and not np.ma.is_masked(dec_deg):
                c = SkyCoord(ra=ra_deg*u.deg, dec=dec_deg*u.deg)
                ra = c.ra.to_string(unit=u.hour, sep=' ', precision=2, pad=True)
                dec = c.dec.to_string(unit=u.deg, sep=' ', precision=2, alwayssign=True, pad=True)
        
        # Get Magnitude (V)
        mag = row['V'] if 'V' in result.colnames and not np.ma.is_masked(row['V']) else None
        
        # Get Object Type
        otype_raw = row['otype'] if 'otype' in result.colnames else None
        otype = OBJECT_TYPE_MAP.get(otype_raw, otype_raw) # Map or fallback to raw

        # 1. Try to get Redshift (z)
        z = None
        if 'rvz_redshift' in result.colnames and not np.ma.is_masked(row['rvz_redshift']):
            z = row['rvz_redshift']
        elif 'rvz_radvel' in result.colnames and not np.ma.is_masked(row['rvz_radvel']):
            v = row['rvz_radvel']
            z = v / c_kms
            
        # 2. Try to get Direct Distance (Parallax Avg OR Most Recent)
        distance_mly = None
        method = "Unknown"
        
        parallax_dists_mpc = []
        other_dists = [] # List of (year, dist_mpc, method)
        
        if 'dist' in result.colnames and 'method' in result.colnames:
            for r in result:
                if np.ma.is_masked(r['dist']):
                    continue
                
                d_val = r['dist']
                d_unit = r['unit']
                
                # Clean unit string
                if isinstance(d_unit, str):
                    d_unit = d_unit.strip()
                
                d_mpc = 0.0
                if d_unit == 'Mpc':
                    d_mpc = d_val
                elif d_unit == 'kpc':
                    d_mpc = d_val * KPC_TO_MPC
                elif d_unit == 'pc':
                    d_mpc = d_val * 1e-6
                else:
                    continue # Skip unknown units

                method_str = r['method']
                if np.ma.is_masked(method_str):
                    method_str = "Unknown"
                
                if isinstance(method_str, str) and method_str.strip() == 'paral':
                    parallax_dists_mpc.append(d_mpc)
                else:
                    # Parse year from bibcode
                    year = 0
                    if 'bibcode' in result.colnames and not np.ma.is_masked(r['bibcode']):
                        bib = r['bibcode']
                        if isinstance(bib, str) and len(bib) >= 4 and bib[:4].isdigit():
                            year = int(bib[:4])
                    other_dists.append((year, d_mpc, method_str))
        
        # Priority 1: Average Parallax
        if len(parallax_dists_mpc) > 0:
            avg_mpc = sum(parallax_dists_mpc) / len(parallax_dists_mpc)
            distance_mly = avg_mpc * MPC_TO_LY * LY_TO_MLY
            distance_pc = avg_mpc * 1e6
            method = "Direct Measurement (Parallax Avg)"
            return z, distance_mly, distance_pc, method, ra, dec, mag, otype
            
        # Priority 2: Most Recent Non-Parallax
        if len(other_dists) > 0:
            # Sort by year descending
            other_dists.sort(key=lambda x: x[0], reverse=True)
            best_match = other_dists[0]
            distance_mly = best_match[1] * MPC_TO_LY * LY_TO_MLY
            distance_pc = best_match[1] * 1e6
            method = f"Direct Measurement (Most Recent: {best_match[0]})"
            return z, distance_mly, distance_pc, method, ra, dec, mag, otype

        # 3. Fallback: Hubble's Law
        if z is not None:
            # Use abs(z) to ensure positive distance
            dist_mpc = (c_kms * abs(z)) / H0
            distance_mly = dist_mpc * MPC_TO_LY * LY_TO_MLY
            distance_pc = dist_mpc * 1e6
            method = "Hubble's Law (Approx)"
            
            return z, distance_mly, distance_pc, method, ra, dec, mag, otype
            
        return None, None, None, "No Data", ra, dec, mag, otype

    except Exception as e:
        print(f"Error querying {object_name}: {e}")
        return None, None, None, "Error", None, None, None, None

def parse_objects_file(filepath):
    objects = []
    try:
        with open(filepath, 'r') as f:
            lines = f.readlines()
        
        for line in lines:
            line = line.strip()
            if not line or ';' not in line:
                continue
            if line.startswith('Name;'):
                continue
                
            parts = line.split(';')
            if len(parts) > 0:
                name = parts[0].strip()
                if name:
                    objects.append(name)
    except FileNotFoundError:
        print(f"Error: File not found at {filepath}")
        return []
    
    return objects

def main():
    parser = argparse.ArgumentParser(description='Calculate astronomical distances and fetch coordinates via Simbad TAP.')
    parser.add_argument('input_file', nargs='?', help='Path to the input text file containing object names')
    parser.add_argument('--output', help='Path to the output Excel file')
    args = parser.parse_args()

    input_file = args.input_file
    
    if input_file is None:
        input_file = input("Enter input filename: ")

    # Determine output file path based on input file location
    if args.output:
        output_file = args.output
    else:
        input_dir = os.path.dirname(os.path.abspath(input_file))
        output_filename = 'Astronomical_Distances_TAP.xlsx'
        output_file = os.path.join(input_dir, output_filename)
    
    print(f"Reading objects from {input_file}...")
    object_names = parse_objects_file(input_file)
    
    if not object_names:
        print("No objects found or file could not be read.")
        return

    print(f"Found {len(object_names)} objects.")
    
    results = []
    
    for name in object_names:
        print(f"Processing {name}...")
        z, dist_mly, dist_pc, method, ra, dec, mag, otype = get_object_data_tap(name)
        
        results.append({
            'Object Name': name,
            'Object Type': otype if otype is not None else 'N/A',
            'RA': ra,
            'Dec': dec,
            'Magnitude': mag if mag is not None else 'N/A',
            'Redshift': z if z is not None else 'N/A',
            'Distance (Parsecs)': dist_pc if dist_pc is not None else 'N/A',
            'Distance (Million Light Years)': dist_mly if dist_mly is not None else 'N/A',
            'Method': method
        })
        # Be nice to the server
        time.sleep(0.1)
        
    df = pd.DataFrame(results)
    df = df[['Object Name', 'Object Type', 'RA', 'Dec', 'Magnitude', 'Redshift', 'Distance (Parsecs)', 'Distance (Million Light Years)', 'Method']]
    
    print(f"Writing results to {output_file}...")
    try:
        # Use ExcelWriter with openpyxl engine to format the output
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Distances')
            
            # Get the workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Distances']
            
            # Freeze the first row
            worksheet.freeze_panes = 'A2'
            
            # Auto-adjust column widths
            for column in df:
                column_length = max(df[column].astype(str).map(len).max(), len(column))
                col_idx = df.columns.get_loc(column)
                writer.sheets['Distances'].column_dimensions[chr(65 + col_idx)].width = column_length + 2
                
        print("Done.")
    except Exception as e:
        print(f"Error writing to Excel: {e}")

if __name__ == "__main__":
    main()
