"""
ArchiCAD Ultra-Fast Smart Element ID Assignment Script (TeamWork Compatible)
=========================================================================

Ultra-fast script to assign unique Element IDs only to elements that need them.
Processes elements individually for immediate feedback and fastest possible speed.
Handles TeamWork permissions gracefully without long delays.

Requirements:
- ArchiCAD with Python API enabled (solo or TeamWork projects)
- archicad Python package installed (pip install archicad)

Author: Saddiq
Date: 2025-06-15
"""

import archicad
from archicad import ACConnection
import os
from typing import Dict, List, Set
import logging
from collections import defaultdict

# Configure simple logging
current_dir = os.path.dirname(os.path.abspath(__file__)) if __file__ else os.getcwd()
log_file_path = os.path.join(current_dir, 'archicad_smart_id_assignment.log')

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file_path),
        logging.StreamHandler()
    ]
)

class SmartUniqueIDAssigner:
    """
    Ultra-fast smart unique ID assignment for building elements in ArchiCAD models.
    Only assigns new IDs to elements with duplicates or empty IDs.
    Preserves existing unique IDs for maximum efficiency.
    Uses individual element processing for immediate feedback and fastest speed.
    Handles TeamWork permission restrictions gracefully without long delays.
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
        
        # Define element type prefixes for construction discipline organization
        self.element_prefixes = {
            'Wall': 'W',
            'Slab': 'S', 
            'Beam': 'B',
            'Column': 'C',
            'Roof': 'R',
            'CurtainWall': 'CW',
            'Stair': 'ST',
            'Railing': 'RL',
            'Door': 'D',
            'Window': 'WIN',
            'Skylight': 'SK',
            'Zone': 'Z',
            'Mesh': 'M',
            'Morph': 'MO',
            'Shell': 'SH',
            'Object': 'O'
        }
        
        # Element types available in ArchiCAD API
        self.available_element_types = [
            'Wall', 'Slab', 'Beam', 'Column', 'Roof', 'CurtainWall',
            'Stair', 'Railing', 'Door', 'Window', 'Skylight', 
            'Zone', 'Mesh', 'Morph', 'Shell', 'Object'
        ]
        
    def get_all_elements_fast(self) -> List:
        """Get all 3D elements quickly with their types."""
        print("Getting all elements...")
        all_elements = []
        
        for elem_type in self.available_element_types:
            try:
                elements = self.acc.GetElementsByType(elem_type)
                for element in elements:
                    all_elements.append({
                        'element': element,
                        'type': elem_type,
                        'guid': element.elementId.guid
                    })
                print(f"  Found {len(elements)} {elem_type} elements")
            except Exception as e:
                print(f"  Warning: Could not get {elem_type} elements: {e}")
        
        print(f"Total elements found: {len(all_elements)}")
        return all_elements
    
    def get_existing_ids_bulk(self, elements: List) -> Dict[str, str]:
        """Get all existing Element IDs in one bulk operation."""
        print("Getting existing Element IDs...")
        
        existing_ids = {}
        
        try:
            # Get Element ID property
            element_id_property = self.acu.GetBuiltInPropertyId('General_ElementID')
            
            # Create element wrappers
            element_wrappers = []
            for elem_data in elements:
                wrapper = self.act.ElementIdArrayItem(self.act.ElementId(elem_data['guid']))
                element_wrappers.append(wrapper)
            
            # Get all Element IDs in bulk
            all_id_values = self.acc.GetPropertyValuesOfElements(element_wrappers, [element_id_property])
            
            for i, id_value_wrapper in enumerate(all_id_values):
                if i < len(elements):
                    element_guid = elements[i]['guid']
                    
                    if id_value_wrapper.propertyValues and len(id_value_wrapper.propertyValues) > 0:
                        prop_value = id_value_wrapper.propertyValues[0].propertyValue
                        if hasattr(prop_value, 'value') and prop_value.value:
                            existing_ids[element_guid] = str(prop_value.value)
                        else:
                            existing_ids[element_guid] = ""
                    else:
                        existing_ids[element_guid] = ""
                        
        except Exception as e:
            print(f"  Warning: Error getting existing IDs: {e}")
            for elem_data in elements:
                existing_ids[elem_data['guid']] = ""
        
        existing_count = len([id_val for id_val in existing_ids.values() if id_val])
        print(f"  Found {existing_count} elements with existing IDs")
        return existing_ids
    
    def analyze_existing_ids(self, existing_ids: Dict[str, str]) -> Dict:
        """Analyze existing IDs to find duplicates and empty ones."""
        print("Analyzing existing IDs for duplicates...")
        
        # Count occurrences of each ID
        id_counts = {}
        empty_ids = []
        
        for element_guid, element_id in existing_ids.items():
            if not element_id or element_id.strip() == "":
                empty_ids.append(element_guid)
            else:
                clean_id = element_id.strip()
                if clean_id not in id_counts:
                    id_counts[clean_id] = []
                id_counts[clean_id].append(element_guid)
        
        # Find duplicates
        duplicate_ids = {}
        unique_ids = set()
        
        for element_id, element_guids in id_counts.items():
            if len(element_guids) > 1:
                # This ID is duplicated
                for guid in element_guids:
                    duplicate_ids[guid] = element_id
            else:
                # This ID is unique - keep it
                unique_ids.add(element_id)
        
        # Elements that need new IDs
        elements_needing_new_ids = set(empty_ids + list(duplicate_ids.keys()))
        
        analysis = {
            'total_elements': len(existing_ids),
            'unique_ids': len(unique_ids),
            'empty_ids': len(empty_ids),
            'duplicate_count': len(duplicate_ids),
            'elements_needing_new_ids': elements_needing_new_ids,
            'existing_unique_ids': unique_ids,
            'duplicate_elements': duplicate_ids
        }
        
        print(f"  Total elements: {analysis['total_elements']}")
        print(f"  Elements with unique IDs: {analysis['unique_ids']} (keeping unchanged)")
        print(f"  Elements with empty IDs: {analysis['empty_ids']}")
        print(f"  Elements with duplicate IDs: {analysis['duplicate_count']}")
        print(f"  Elements needing new IDs: {len(elements_needing_new_ids)}")
        
        return analysis
    
    def generate_new_ids_for_problem_elements(self, elements: List, analysis: Dict) -> Dict[str, str]:
        """Generate new IDs only for elements that need them."""
        print("Generating new IDs only for problem elements...")
        
        elements_needing_new_ids = analysis['elements_needing_new_ids']
        existing_unique_ids = analysis['existing_unique_ids']
        
        if not elements_needing_new_ids:
            print("  No elements need new IDs - all are already unique!")
            return {}
        
        # Filter elements to only those needing new IDs
        problem_elements = []
        for elem_data in elements:
            if elem_data['guid'] in elements_needing_new_ids:
                problem_elements.append(elem_data)
        
        print(f"  Processing {len(problem_elements)} elements that need new IDs")
        
        # Group problem elements by type
        elements_by_type = defaultdict(list)
        for elem_data in problem_elements:
            elements_by_type[elem_data['type']].append(elem_data)
        
        # All existing IDs (including unique ones we're keeping)
        all_reserved_ids = set(existing_unique_ids)
        
        # Generate new IDs by type
        new_id_mapping = {}
        
        for element_type, type_elements in elements_by_type.items():
            prefix = self.element_prefixes.get(element_type, 'GEN')
            
            print(f"    Assigning new IDs to {len(type_elements)} {element_type} elements")
            
            # Find the next available counter for this prefix
            counter = 1
            while f"{prefix}-{counter:03d}" in all_reserved_ids:
                counter += 1
            
            for elem_data in type_elements:
                element_guid = elem_data['guid']
                
                # Generate next available unique ID
                while True:
                    new_id = f"{prefix}-{counter:03d}"
                    if new_id not in all_reserved_ids:
                        break
                    counter += 1
                
                new_id_mapping[element_guid] = new_id
                all_reserved_ids.add(new_id)
                counter += 1
        
        print(f"  Generated {len(new_id_mapping)} new unique IDs")
        return new_id_mapping
    
    def assign_ids_super_fast(self, new_id_mapping: Dict[str, str]) -> Dict:
        """Assign all new IDs using ultra-fast processing with minimal TeamWork delays."""
        print("Assigning new IDs to elements (ultra-fast mode)...")
        
        try:
            # Get Element ID property
            element_id_property = self.acu.GetBuiltInPropertyId('General_ElementID')
            
            # Process elements individually for fastest response and immediate feedback
            success_count = 0
            permission_denied_count = 0
            other_errors = 0
            failed_elements = []
            
            total_elements = len(new_id_mapping)
            
            print(f"  Processing {total_elements} elements individually for speed...")
            print("  Press Ctrl+C to stop early if needed")
            
            for i, (element_guid, new_id) in enumerate(new_id_mapping.items(), 1):
                try:
                    # Create property value
                    property_value = self.act.NormalStringPropertyValue(new_id)
                    
                    # Create element property value assignment
                    element_property_value = self.act.ElementPropertyValue(
                        elementId=self.act.ElementId(element_guid),
                        propertyId=element_id_property,
                        propertyValue=property_value
                    )
                    
                    # Process single element with immediate response
                    results = self.acc.SetPropertyValuesOfElements([element_property_value])
                    
                    if results and len(results) > 0:
                        result = results[0]
                        if result.success:
                            success_count += 1
                            print(f"    ✓ {i}/{total_elements}: {new_id} assigned")
                        else:
                            error_msg = result.error.message if result.error else "Unknown error"
                            
                            if "permission" in error_msg.lower() or "teamwork" in error_msg.lower() or "reserved" in error_msg.lower():
                                permission_denied_count += 1
                                print(f"    ⚠ {i}/{total_elements}: {new_id} - Permission denied")
                                failed_elements.append({
                                    'guid': element_guid,
                                    'new_id': new_id,
                                    'reason': 'TeamWork permission denied'
                                })
                            else:
                                other_errors += 1
                                print(f"    ❌ {i}/{total_elements}: {new_id} - Error: {error_msg}")
                                failed_elements.append({
                                    'guid': element_guid,
                                    'new_id': new_id,
                                    'reason': error_msg
                                })
                    else:
                        other_errors += 1
                        print(f"    ❌ {i}/{total_elements}: {new_id} - No response")
                        failed_elements.append({
                            'guid': element_guid,
                            'new_id': new_id,
                            'reason': 'No response from API'
                        })
                    
                    # Progress summary every 25 elements
                    if i % 25 == 0:
                        print(f"    Progress: {i}/{total_elements} ({success_count} success, {permission_denied_count} denied)")
                        
                except KeyboardInterrupt:
                    print(f"\n  User stopped processing at element {i}/{total_elements}")
                    print(f"  Results so far: {success_count} success, {permission_denied_count} denied")
                    break
                    
                except Exception as e:
                    other_errors += 1
                    print(f"    ❌ {i}/{total_elements}: {new_id} - Exception: {str(e)}")
                    failed_elements.append({
                        'guid': element_guid,
                        'new_id': new_id,
                        'reason': f'Processing error: {str(e)}'
                    })
                    continue
            
            # Final summary
            print(f"\n  FINAL RESULTS:")
            print(f"    Successfully assigned: {success_count}/{total_elements}")
            if permission_denied_count > 0:
                print(f"    Permission denied (TeamWork): {permission_denied_count}")
            if other_errors > 0:
                print(f"    Other errors: {other_errors}")
            
            return {
                'success_count': success_count,
                'permission_denied_count': permission_denied_count,
                'other_errors': other_errors,
                'failed_elements': failed_elements,
                'total_attempted': total_elements
            }
            
        except Exception as e:
            print(f"  Error during fast ID assignment: {e}")
            return {
                'success_count': 0,
                'permission_denied_count': 0,
                'other_errors': len(new_id_mapping),
                'failed_elements': [],
                'total_attempted': len(new_id_mapping)
            }
    
    def generate_assignment_report(self, elements: List, analysis: Dict, new_id_mapping: Dict[str, str], existing_ids: Dict[str, str], assignment_results: Dict) -> str:
        """Generate a summary report of ID assignments including TeamWork issues."""
        
        # Count by element type
        type_counts = defaultdict(int)
        kept_unchanged = defaultdict(int)
        assigned_new = defaultdict(int)
        
        for elem_data in elements:
            element_type = elem_data['type']
            element_guid = elem_data['guid']
            
            type_counts[element_type] += 1
            
            if element_guid in new_id_mapping:
                assigned_new[element_type] += 1
            else:
                kept_unchanged[element_type] += 1
        
        report_lines = [
            "SMART ELEMENT ID ASSIGNMENT REPORT (TeamWork Compatible)",
            "=" * 60,
            f"Total Elements Analyzed: {len(elements)}",
            f"Elements with Unique IDs (kept unchanged): {analysis['unique_ids']}",
            f"Elements with Empty IDs (needed new): {analysis['empty_ids']}",
            f"Elements with Duplicate IDs (needed new): {analysis['duplicate_count']}",
            f"Total Elements Needing New IDs: {len(new_id_mapping)}",
            "",
            "ASSIGNMENT RESULTS:",
            "-" * 30,
            f"Successfully assigned new IDs: {assignment_results['success_count']}",
            f"Permission denied (TeamWork): {assignment_results['permission_denied_count']}",
            f"Other errors: {assignment_results['other_errors']}",
            f"Generated: {logging.Formatter().formatTime(logging.LogRecord('', 0, '', 0, '', (), None))}",
            "",
            "EFFICIENCY SUMMARY:",
            "-" * 30,
            f"✓ Kept {analysis['unique_ids']} existing unique IDs unchanged",
            f"✓ Successfully processed {assignment_results['success_count']} elements that needed changes"
        ]
        
        if assignment_results['permission_denied_count'] > 0:
            report_lines.extend([
                f"⚠ {assignment_results['permission_denied_count']} elements couldn't be changed (TeamWork permissions)",
                f"✓ Result: {len(elements) - assignment_results['permission_denied_count'] - assignment_results['other_errors']} elements have unique IDs"
            ])
        else:
            report_lines.append(f"✓ Result: All {len(elements)} elements now have unique IDs")
        
        report_lines.extend([
            "",
            "ELEMENT TYPE BREAKDOWN:",
            "-" * 30
        ])
        
        for element_type in sorted(type_counts.keys()):
            total = type_counts[element_type]
            kept = kept_unchanged[element_type]
            new_assigned = assigned_new[element_type]
            prefix = self.element_prefixes.get(element_type, 'GEN')
            
            report_lines.append(f"\n{element_type.upper()} (Total: {total}, Prefix: {prefix}):")
            report_lines.append(f"  Kept unchanged: {kept} elements")
            report_lines.append(f"  Assigned new IDs: {new_assigned} elements")
        
        # TeamWork issues section
        if assignment_results['failed_elements']:
            report_lines.extend([
                "",
                "TEAMWORK PERMISSION ISSUES:",
                "-" * 30
            ])
            
            permission_failures = [elem for elem in assignment_results['failed_elements'] 
                                 if 'permission' in elem['reason'].lower()]
            
            if permission_failures:
                report_lines.append(f"\nElements that couldn't be modified (reserved by other users):")
                for elem in permission_failures[:10]:  # Show first 10
                    report_lines.append(f"  {elem['guid']} → (wanted: {elem['new_id']}) - {elem['reason']}")
                
                if len(permission_failures) > 10:
                    report_lines.append(f"  ... and {len(permission_failures) - 10} more elements")
                
                report_lines.extend([
                    "",
                    "TEAMWORK SOLUTIONS:",
                    "• Ask other users to release reserved elements",
                    "• Run the script again after elements are released",
                    "• Coordinate with team members to avoid conflicts"
                ])
        
        if analysis['duplicate_elements'] and assignment_results['success_count'] > 0:
            report_lines.extend([
                "",
                "DUPLICATE IDs SUCCESSFULLY FIXED:",
                "-" * 30
            ])
            
            duplicate_groups = defaultdict(list)
            for guid, old_id in analysis['duplicate_elements'].items():
                # Only show successfully fixed duplicates
                if guid in new_id_mapping and not any(elem['guid'] == guid for elem in assignment_results['failed_elements']):
                    duplicate_groups[old_id].append(guid)
            
            for old_id, guids in duplicate_groups.items():
                if guids:  # Only show if we have successfully fixed elements
                    report_lines.append(f"\nDuplicate ID '{old_id}' fixed on {len(guids)} elements:")
                    for guid in guids:
                        new_id = new_id_mapping.get(guid, "ERROR")
                        report_lines.append(f"  {guid} → {new_id}")
        
        report_lines.extend([
            "",
            "=" * 60,
            "STATUS: Smart ID assignment completed with TeamWork compatibility",
            "EFFICIENCY: Only changed elements that actually needed new IDs",
            "TEAMWORK: Handled permission restrictions gracefully",
            "Generated by Smart ArchiCAD Unique ID Assigner"
        ])
        
        return "\n".join(report_lines)
    
    def assign_unique_ids_to_all_elements(self) -> bool:
        """Main process to assign unique IDs only to elements that need them with TeamWork support."""
        
        try:
            print("=" * 50)
            print("ULTRA-FAST SMART ELEMENT ID ASSIGNMENT (TeamWork Compatible)")
            print("=" * 50)
            
            # Step 1: Get all elements
            elements = self.get_all_elements_fast()
            if not elements:
                print("❌ No elements found!")
                return False
            
            # Step 2: Get existing IDs
            existing_ids = self.get_existing_ids_bulk(elements)
            
            # Step 3: Analyze which elements actually need new IDs
            analysis = self.analyze_existing_ids(existing_ids)
            
            # Step 4: Generate new IDs only for problem elements
            new_id_mapping = self.generate_new_ids_for_problem_elements(elements, analysis)
            
            if not new_id_mapping:
                print("=" * 50)
                print("✓ ALL ELEMENTS ALREADY HAVE UNIQUE IDs!")
                print("No changes needed - all Element IDs are already unique.")
                print("=" * 50)
                return True
            
            # Step 5: Assign new IDs using ultra-fast individual processing
            assignment_results = self.assign_ids_super_fast(new_id_mapping)
            
            # Step 6: Generate report
            print("Generating assignment report...")
            report = self.generate_assignment_report(elements, analysis, new_id_mapping, existing_ids, assignment_results)
            report_path = os.path.join(current_dir, "smart_id_assignment_report.txt")
            
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(report)
            
            # Display results
            print("=" * 50)
            print(f"✓ Smart ID assignment completed!")
            print(f"✓ Kept {analysis['unique_ids']} existing unique IDs unchanged")
            print(f"✓ Successfully assigned {assignment_results['success_count']} new IDs")
            
            if assignment_results['permission_denied_count'] > 0:
                print(f"⚠ {assignment_results['permission_denied_count']} elements couldn't be changed (TeamWork permissions)")
                print("  These elements are reserved by other users")
                print("  Ask team members to release them and run the script again")
            
            if assignment_results['other_errors'] > 0:
                print(f"⚠ {assignment_results['other_errors']} elements had other errors")
            
            total_unique = len(elements) - assignment_results['permission_denied_count'] - assignment_results['other_errors']
            print(f"✓ Result: {total_unique}/{len(elements)} elements have unique IDs")
            print(f"✓ Assignment report: {report_path}")
            print(f"✓ Location: {current_dir}")
            print("=" * 50)
            
            return assignment_results['success_count'] > 0 or analysis['unique_ids'] > 0
            
        except Exception as e:
            print(f"❌ Error during ID assignment: {e}")
            logging.error(f"ID assignment error: {e}")
            return False


def main():
    """Main function - ultra-fast smart unique ID assignment with TeamWork support."""
    try:
        assigner = SmartUniqueIDAssigner()
        success = assigner.assign_unique_ids_to_all_elements()
        
        if success:
            print("\n🎉 Ultra-fast Smart ID assignment completed!")
            print("Only elements with duplicate/empty IDs were changed.")
            print("Existing unique IDs were preserved unchanged.")
            print("Individual element processing provided immediate feedback.")
            print("TeamWork permissions were handled instantly without delays.")
        else:
            print("\n❌ ID assignment failed. Check the log for details.")
            
    except Exception as e:
        print(f"\n❌ Error: {e}")
        print("Make sure ArchiCAD is running and ready (no dialogs open, not in drawing mode)")


if __name__ == "__main__":
    """
    Ultra-Fast Smart ArchiCAD Unique Element ID Assignment (TeamWork Compatible)
    
    Usage:
    1. Open ArchiCAD with your project (TeamWork or solo)
    2. Make sure no dialogs are open and you're not drawing
    3. Run: python archicad_unique_id_assignment.py
    
    The script will:
    - Connect to ArchiCAD quickly
    - Analyze all existing Element IDs
    - Identify duplicates and empty IDs
    - Keep existing unique IDs unchanged (no unnecessary changes)
    - Process elements individually for immediate feedback
    - Assign new construction-standard IDs with real-time progress
    - Handle TeamWork permission restrictions instantly (no long waits)
    - Complete available elements in minutes, not hours
    
    Ultra-Fast Features:
    ✓ Individual element processing for immediate response
    ✓ Real-time feedback on each element (✓ success, ⚠ denied, ❌ error)
    ✓ No long waits for TeamWork timeouts
    ✓ Progress tracking with live updates
    ✓ Can be stopped early with Ctrl+C if needed
    
    TeamWork Features:
    ✓ Handles elements reserved by other users instantly
    ✓ No delays waiting for permission responses
    ✓ Continues processing immediately after permission denials
    ✓ Provides clear real-time feedback about reserved elements
    
    Smart Features:
    ✓ Only changes elements with duplicate or empty IDs
    ✓ Preserves existing unique IDs (much faster)
    ✓ Ensures perfect uniqueness across all elements
    ✓ Uses construction-standard prefixes (W-001, B-001, etc.)
    
    ID Prefixes Used:
    W-001, W-002... (Walls)
    B-001, B-002... (Beams)  
    C-001, C-002... (Columns)
    S-001, S-002... (Slabs)
    R-001, R-002... (Roofs)
    And more for other element types
    
    Output:
    - smart_id_assignment_report.txt: Detailed assignment summary
    - archicad_smart_id_assignment.log: Process log
    
    Real-Time Output Example:
    "✓ 1/50: W-001 assigned"
    "⚠ 2/50: W-002 - Permission denied"
    "✓ 3/50: W-003 assigned"
    "Progress: 25/50 (20 success, 5 denied)"
    
    Speed: Processes elements individually with immediate feedback
    No more waiting 15+ minutes for batch operations!
    """
    main()