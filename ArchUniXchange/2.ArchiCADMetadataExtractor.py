"""
ArchiCAD Fast Comprehensive Element Data Extractor
=================================================

Fast script to extract ALL element data to CSV in under 2 minutes.
Gets all the information you see in ArchiCAD's element properties palette including:
- Element ID, Type, GUID, Layer, Home Story
- Position, Renovation Status, Structural Function  
- All geometric measurements (Length, Area, Volume, etc.)
- All custom properties and classifications
- Everything visible in the properties palette

Requirements:
- ArchiCAD with Python API enabled
- archicad Python package installed (pip install archicad)

Author: Saddiq
Date: 2025-06-15
"""

import archicad
from archicad import ACConnection
import csv
import os
from typing import Dict, List
import logging
from datetime import datetime

# Configure simple logging
current_dir = os.path.dirname(os.path.abspath(__file__)) if __file__ else os.getcwd()
log_file_path = os.path.join(current_dir, 'archicad_fast_extraction.log')

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file_path),
        logging.StreamHandler()
    ]
)

class FastElementExtractor:
    """
    Fast element data extractor for ArchiCAD.
    Gets ALL properties and measurements you see in the element properties palette.
    """
    
    def __init__(self):
        """Initialize connection to ArchiCAD."""
        try:
            print("Connecting to ArchiCAD...")
            self.conn = ACConnection.connect()
            self.acc = self.conn.commands
            self.act = self.conn.types
            self.acu = self.conn.utilities
            print("✓ Connected successfully!")
        except Exception as e:
            print(f"❌ Connection failed: {e}")
            print("\nMake sure ArchiCAD is running and ready (no dialogs, not drawing, etc.)")
            raise
        
    def get_all_elements_fast(self) -> List:
        """Get all 3D elements quickly using single API calls."""
        print("Getting all elements...")
        all_elements = []
        
        element_types = ['Wall', 'Slab', 'Beam', 'Column', 'Roof', 'CurtainWall',
                        'Stair', 'Railing', 'Door', 'Window', 'Skylight', 
                        'Zone', 'Mesh', 'Morph', 'Shell', 'Object']
        
        for elem_type in element_types:
            try:
                elements = self.acc.GetElementsByType(elem_type)
                for element in elements:
                    all_elements.append({
                        'element': element,
                        'type': elem_type
                    })
                print(f"  Found {len(elements)} {elem_type} elements")
            except Exception as e:
                print(f"  Warning: Could not get {elem_type} elements: {e}")
        
        print(f"Total elements found: {len(all_elements)}")
        return all_elements
    
    def get_all_properties_bulk(self, elements: List) -> Dict:
        """Get ALL available properties for all elements in bulk operations."""
        print("Getting all properties and measurements...")
        
        # Create element wrappers
        element_wrappers = []
        element_map = {}
        
        for i, elem_data in enumerate(elements):
            wrapper = self.act.ElementIdArrayItem(self.act.ElementId(elem_data['element'].elementId.guid))
            element_wrappers.append(wrapper)
            element_map[i] = elem_data
        
        # Get ALL property IDs available in the project
        print("  Discovering all available properties...")
        try:
            all_property_ids = self.acc.GetAllPropertyIds()
            property_details = self.acc.GetDetailsOfProperties(all_property_ids)
            print(f"  Found {len(all_property_ids)} properties")
        except Exception as e:
            print(f"  Error getting property definitions: {e}")
            return {}
        
        # Get all property values in one bulk operation
        results = {}
        try:
            print("  Extracting property values for all elements...")
            all_prop_values = self.acc.GetPropertyValuesOfElements(element_wrappers, all_property_ids)
            
            for i, prop_values_wrapper in enumerate(all_prop_values):
                if i >= len(element_map):
                    continue
                    
                element_guid = element_map[i]['element'].elementId.guid
                element_type = element_map[i]['type']
                
                results[element_guid] = {
                    'Element_Type': element_type,
                    'Element_GUID': element_guid
                }
                
                # Extract all property values
                if prop_values_wrapper.propertyValues:
                    for j, prop_value in enumerate(prop_values_wrapper.propertyValues):
                        if j < len(property_details):
                            try:
                                prop_name = property_details[j].propertyDefinition.name
                                
                                if hasattr(prop_value.propertyValue, 'value') and prop_value.propertyValue.value is not None:
                                    # Handle different value types
                                    value = prop_value.propertyValue.value
                                    if isinstance(value, (int, float)):
                                        results[element_guid][prop_name] = str(value)
                                    else:
                                        results[element_guid][prop_name] = str(value)
                                else:
                                    results[element_guid][prop_name] = "<Undefined>"
                                    
                            except Exception as e:
                                # Skip problematic properties but continue processing
                                continue
                
                # Progress indicator
                if (i + 1) % 50 == 0:
                    print(f"    Processed {i + 1}/{len(element_map)} elements...")
                    
        except Exception as e:
            print(f"  Error getting bulk properties: {e}")
            # Fallback to basic data only
            for i, elem_data in enumerate(elements):
                element_guid = elem_data['element'].elementId.guid
                results[element_guid] = {
                    'Element_Type': elem_data['type'],
                    'Element_GUID': element_guid
                }
        
        print(f"  ✓ Extracted properties for {len(results)} elements")
        return results
    
    def get_classifications_fast(self, elements: List) -> Dict:
        """Get classifications quickly."""
        print("Getting classifications...")
        
        classifications = {}
        
        try:
            # Get classification systems
            class_systems = self.acc.GetAllClassificationSystems()
            if not class_systems:
                print("  No classification systems found")
                return {}
            
            # Create element wrappers
            element_wrappers = []
            for elem_data in elements:
                wrapper = self.act.ElementIdArrayItem(self.act.ElementId(elem_data['element'].elementId.guid))
                element_wrappers.append(wrapper)
            
            # Create classification system IDs
            system_ids = []
            system_names = {}
            for system in class_systems:
                system_id = self.act.ClassificationSystemIdArrayItem(system.classificationSystemId)
                system_ids.append(system_id)
                system_names[system.classificationSystemId.guid] = system.name
            
            # Get all classifications in bulk
            all_classifications = self.acc.GetClassificationsOfElements(element_wrappers, system_ids)
            
            for i, class_wrapper in enumerate(all_classifications):
                element_guid = elements[i]['element'].elementId.guid
                classifications[element_guid] = {}
                
                if class_wrapper.classificationIds:
                    for classification in class_wrapper.classificationIds:
                        try:
                            if hasattr(classification, 'classificationId') and classification.classificationId:
                                system_guid = classification.classificationId.classificationSystemId.guid
                                system_name = system_names.get(system_guid, 'Unknown_System')
                                
                                if (hasattr(classification.classificationId, 'classificationItemId') and 
                                    classification.classificationId.classificationItemId):
                                    item_name = classification.classificationId.classificationItemId.name
                                else:
                                    item_name = ""
                                
                                # Clean system name for CSV
                                clean_name = system_name.replace(' ', '_').replace('-', '_').replace('.', '_')
                                classifications[element_guid][f'Classification_{clean_name}'] = item_name
                        except Exception as e:
                            continue
                            
        except Exception as e:
            print(f"Warning: Error getting classifications: {e}")
            for elem_data in elements:
                classifications[elem_data['element'].elementId.guid] = {}
        
        return classifications
    
    def extract_to_csv(self, filename: str = None) -> bool:
        """Extract all element data to CSV quickly."""
        try:
            if filename is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"archicad_elements_{timestamp}.csv"
            
            csv_path = os.path.join(current_dir, filename)
            
            print("=" * 50)
            print("FAST COMPREHENSIVE ARCHICAD ELEMENT EXTRACTION")
            print("=" * 50)
            
            # Step 1: Get all elements (fast)
            elements = self.get_all_elements_fast()
            if not elements:
                print("❌ No elements found!")
                return False
            
            # Step 2: Get all properties and measurements (bulk operation)
            properties_data = self.get_all_properties_bulk(elements)
            
            # Step 3: Get classifications (bulk operation)  
            classifications_data = self.get_classifications_fast(elements)
            
            # Step 4: Combine data and write CSV
            print("Writing CSV file...")
            
            # Determine all column names
            all_columns = set(['Element_GUID', 'Element_Type'])
            for guid, data in properties_data.items():
                all_columns.update(data.keys())
            for guid, data in classifications_data.items():
                all_columns.update(data.keys())
            
            # Sort columns for consistent output
            sorted_columns = ['Element_GUID', 'Element_Type'] + sorted([col for col in all_columns if col not in ['Element_GUID', 'Element_Type']])
            
            # Write CSV
            with open(csv_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=sorted_columns)
                writer.writeheader()
                
                for elem_data in elements:
                    guid = elem_data['element'].elementId.guid
                    
                    # Combine all data for this element
                    row_data = properties_data.get(guid, {})
                    row_data.update(classifications_data.get(guid, {}))
                    
                    # Ensure all columns are present
                    complete_row = {}
                    for col in sorted_columns:
                        complete_row[col] = row_data.get(col, "")
                    
                    writer.writerow(complete_row)
            
            print("=" * 50)
            print(f"✓ SUCCESS! Extracted {len(elements)} elements")
            print(f"✓ CSV saved: {filename}")
            print(f"✓ Properties extracted: {len(sorted_columns)}")
            print(f"✓ Includes: Element details, measurements, classifications")
            print(f"✓ Location: {current_dir}")
            print("=" * 50)
            
            return True
            
        except Exception as e:
            print(f"❌ Error during extraction: {e}")
            logging.error(f"Extraction error: {e}")
            return False


def main():
    """Main function - simple and fast."""
    try:
        extractor = FastElementExtractor()
        success = extractor.extract_to_csv()
        
        if success:
            print("\n🎉 Extraction completed successfully!")
        else:
            print("\n❌ Extraction failed. Check the log for details.")
            
    except Exception as e:
        print(f"\n❌ Error: {e}")
        print("Make sure ArchiCAD is running and ready (no dialogs open, not in drawing mode)")


if __name__ == "__main__":
    """
    Fast ArchiCAD Comprehensive Element Data Extractor
    
    Usage:
    1. Open ArchiCAD with your project
    2. Make sure no dialogs are open and you're not drawing
    3. Run: python archicad_metadata_extractor.py
    
    The script will:
    - Connect to ArchiCAD quickly
    - Get all 3D elements in one go
    - Extract ALL properties and measurements including:
      * Basic info (Element ID, Type, GUID, Layer, Home Story)
      * Position, Renovation Status, Structural Function
      * All geometric measurements (Length, Area, Volume, etc.)
      * All custom properties (내화 등급, 가연성, 열관류율, etc.)
      * Classifications (ARCHICAD 분류 v 2.0, etc.)
    - Export everything to CSV in under 2 minutes
    
    Output:
    - archicad_elements_[timestamp].csv with ALL element data
    - Everything you see in ArchiCAD's element properties palette
    """
    main()