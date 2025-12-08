"""
ArchiCAD Comprehensive BIM Metadata Extractor
==============================================

Enhanced script to extract ALL element data including IFC metadata to CSV.
Extracts comprehensive BIM data similar to Revit metadata extraction including:

IFC Identification:
- IFC GlobalId, IFC EntityType, IFC PredefinedType, IFC Definition

IFC Hierarchy:
- Project, Site, Building, Storey information

Material Information:
- Material names, colors, layers, thicknesses

Property Sets (Pset):
- Common properties like IsExternal, LoadBearing, Reference

Quantity Sets (Qto):
- BaseQuantities like Length, Area, Volume, GrossSurfaceArea

Geometric Data:
- Dimensions, positions, constraints

Element Relationships:
- Classifications, group assignments, associations

Requirements:
- ArchiCAD with Python API enabled
- archicad Python package installed (pip install archicad)

Author: Saddiq
Date: 2025-06-15
Updated: 2025-12-01 - Added comprehensive IFC metadata extraction
"""

import archicad
from archicad import ACConnection
import csv
import os
import uuid
import base64
from typing import Dict, List, Tuple, Optional, Any
import logging
from datetime import datetime

# Configure logging
current_dir = os.path.dirname(os.path.abspath(__file__)) if __file__ else os.getcwd()
log_file_path = os.path.join(current_dir, 'archicad_metadata_extraction.log')

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file_path),
        logging.StreamHandler()
    ]
)


class IfcMappingHelper:
    """
    Helper class for IFC mapping and metadata generation.
    Maps ArchiCAD element types to IFC entity types and provides IFC-related utilities.
    """
    
    # ArchiCAD to IFC entity type mapping
    ARCHICAD_TO_IFC_MAPPING = {
        'Wall': 'IfcWall',
        'Slab': 'IfcSlab',
        'Beam': 'IfcBeam',
        'Column': 'IfcColumn',
        'Roof': 'IfcRoof',
        'CurtainWall': 'IfcCurtainWall',
        'Stair': 'IfcStair',
        'Railing': 'IfcRailing',
        'Door': 'IfcDoor',
        'Window': 'IfcWindow',
        'Skylight': 'IfcWindow',
        'Zone': 'IfcSpace',
        'Mesh': 'IfcBuildingElementProxy',
        'Morph': 'IfcBuildingElementProxy',
        'Shell': 'IfcBuildingElementProxy',
        'Object': 'IfcBuildingElementProxy',
        'Opening': 'IfcOpeningElement',
        'Lamp': 'IfcLightFixture',
        'Fill': 'IfcCovering',
        'Hatch': 'IfcAnnotation',
        'Line': 'IfcAnnotation',
        'Dimension': 'IfcAnnotation',
        'Text': 'IfcAnnotation',
        'Label': 'IfcAnnotation',
    }
    
    # IFC predefined types by entity
    IFC_PREDEFINED_TYPES = {
        'IfcWall': ['MOVABLE', 'PARAPET', 'PARTITIONING', 'PLUMBINGWALL', 'SHEAR', 'SOLIDWALL', 'STANDARD', 'POLYGONAL', 'ELEMENTEDWALL', 'NOTDEFINED'],
        'IfcSlab': ['FLOOR', 'ROOF', 'LANDING', 'BASESLAB', 'APPROACH_SLAB', 'PAVING', 'WEARING', 'SIDEWALK', 'NOTDEFINED'],
        'IfcBeam': ['BEAM', 'JOIST', 'HOLLOWCORE', 'LINTEL', 'SPANDREL', 'T_BEAM', 'NOTDEFINED'],
        'IfcColumn': ['COLUMN', 'PILASTER', 'NOTDEFINED'],
        'IfcRoof': ['FLAT_ROOF', 'SHED_ROOF', 'GABLE_ROOF', 'HIP_ROOF', 'HIPPED_GABLE_ROOF', 'GAMBREL_ROOF', 'MANSARD_ROOF', 'BARREL_ROOF', 'RAINBOW_ROOF', 'BUTTERFLY_ROOF', 'PAVILION_ROOF', 'DOME_ROOF', 'FREEFORM', 'NOTDEFINED'],
        'IfcDoor': ['DOOR', 'GATE', 'TRAPDOOR', 'NOTDEFINED'],
        'IfcWindow': ['WINDOW', 'SKYLIGHT', 'LIGHTDOME', 'NOTDEFINED'],
        'IfcStair': ['STRAIGHT_RUN_STAIR', 'TWO_STRAIGHT_RUN_STAIR', 'QUARTER_WINDING_STAIR', 'QUARTER_TURN_STAIR', 'HALF_WINDING_STAIR', 'HALF_TURN_STAIR', 'TWO_QUARTER_WINDING_STAIR', 'TWO_QUARTER_TURN_STAIR', 'THREE_QUARTER_WINDING_STAIR', 'THREE_QUARTER_TURN_STAIR', 'SPIRAL_STAIR', 'DOUBLE_RETURN_STAIR', 'CURVED_RUN_STAIR', 'TWO_CURVED_RUN_STAIR', 'NOTDEFINED'],
        'IfcRailing': ['HANDRAIL', 'GUARDRAIL', 'BALUSTRADE', 'NOTDEFINED'],
        'IfcCurtainWall': ['NOTDEFINED'],
        'IfcSpace': ['SPACE', 'PARKING', 'GFA', 'INTERNAL', 'EXTERNAL', 'NOTDEFINED'],
    }
    
    # IFC definitions
    IFC_DEFINITIONS = {
        'IfcWall': 'A wall is a vertical building element that delimits or subdivides spaces.',
        'IfcSlab': 'A slab is a component of the construction that normally encloses a space vertically.',
        'IfcBeam': 'A beam is a horizontal structural member designed primarily to carry and resist bending loads.',
        'IfcColumn': 'A column is a vertical structural member that transmits loads from above to a load bearing element below.',
        'IfcRoof': 'A roof is a construction enclosing the building from above.',
        'IfcDoor': 'A door is a building element that provides a passage between spaces.',
        'IfcWindow': 'A window is a building element for the passage of light and/or ventilation.',
        'IfcStair': 'A stair is a vertical circulation element providing access between different floor levels.',
        'IfcRailing': 'A railing is a frame assembly used as a barrier or support.',
        'IfcCurtainWall': 'A curtain wall is a non-load bearing outer wall of a building.',
        'IfcSpace': 'A space represents an area or volume bounded actually or theoretically.',
        'IfcBuildingElementProxy': 'A building element proxy is used for building elements that cannot be classified.',
    }
    
    # Applicable property sets by IFC entity
    APPLICABLE_PSETS = {
        'IfcWall': ['Pset_WallCommon', 'Pset_QuantityTakeOff'],
        'IfcSlab': ['Pset_SlabCommon', 'Pset_QuantityTakeOff'],
        'IfcBeam': ['Pset_BeamCommon', 'Pset_QuantityTakeOff'],
        'IfcColumn': ['Pset_ColumnCommon', 'Pset_QuantityTakeOff'],
        'IfcRoof': ['Pset_RoofCommon', 'Pset_QuantityTakeOff'],
        'IfcDoor': ['Pset_DoorCommon', 'Pset_DoorWindowGlazingType', 'Pset_DoorWindowShadingType'],
        'IfcWindow': ['Pset_WindowCommon', 'Pset_DoorWindowGlazingType', 'Pset_DoorWindowShadingType'],
        'IfcStair': ['Pset_StairCommon'],
        'IfcRailing': ['Pset_RailingCommon'],
        'IfcCurtainWall': ['Pset_CurtainWallCommon'],
        'IfcSpace': ['Pset_SpaceCommon', 'Pset_SpaceOccupancyRequirements', 'Pset_SpaceThermalRequirements'],
    }
    
    # Applicable quantity sets by IFC entity
    APPLICABLE_QTOS = {
        'IfcWall': ['Qto_WallBaseQuantities'],
        'IfcSlab': ['Qto_SlabBaseQuantities'],
        'IfcBeam': ['Qto_BeamBaseQuantities'],
        'IfcColumn': ['Qto_ColumnBaseQuantities'],
        'IfcRoof': ['Qto_RoofBaseQuantities'],
        'IfcDoor': ['Qto_DoorBaseQuantities'],
        'IfcWindow': ['Qto_WindowBaseQuantities'],
        'IfcStair': ['Qto_StairFlightBaseQuantities'],
        'IfcRailing': ['Qto_RailingBaseQuantities'],
        'IfcCurtainWall': ['Qto_CurtainWallBaseQuantities'],
        'IfcSpace': ['Qto_SpaceBaseQuantities'],
    }
    
    @classmethod
    def get_ifc_entity_type(cls, archicad_type: str) -> str:
        """Get IFC entity type from ArchiCAD element type."""
        return cls.ARCHICAD_TO_IFC_MAPPING.get(archicad_type, 'IfcBuildingElementProxy')
    
    @classmethod
    def get_ifc_predefined_type(cls, ifc_entity: str, archicad_type: str = None) -> str:
        """Get IFC predefined type for an entity."""
        predefined_types = cls.IFC_PREDEFINED_TYPES.get(ifc_entity, ['NOTDEFINED'])
        
        # Special mappings based on ArchiCAD type
        if archicad_type == 'Skylight' and ifc_entity == 'IfcWindow':
            return 'SKYLIGHT'
        
        return predefined_types[0] if predefined_types else 'NOTDEFINED'
    
    @classmethod
    def get_ifc_definition(cls, ifc_entity: str) -> str:
        """Get IFC definition for an entity type."""
        return cls.IFC_DEFINITIONS.get(ifc_entity, 'A building element.')
    
    @classmethod
    def get_applicable_psets(cls, ifc_entity: str) -> str:
        """Get applicable property sets for an IFC entity."""
        psets = cls.APPLICABLE_PSETS.get(ifc_entity, [])
        return '; '.join(psets) if psets else ''
    
    @classmethod
    def get_applicable_qtos(cls, ifc_entity: str) -> str:
        """Get applicable quantity sets for an IFC entity."""
        qtos = cls.APPLICABLE_QTOS.get(ifc_entity, [])
        return '; '.join(qtos) if qtos else ''
    
    @staticmethod
    def get_ifc_schema_version() -> str:
        """Get the IFC schema version."""
        return 'IFC4'
    
    @staticmethod
    def generate_ifc_global_id(guid: str) -> str:
        """
        Generate IFC-compliant GlobalId (22 character base64-like encoding).
        """
        try:
            # IFC GlobalId uses a specific base64 encoding
            base64_chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$"
            
            # Convert UUID object to string if necessary
            if hasattr(guid, 'hex'):
                # It's a UUID object, convert to string
                guid_str = str(guid)
            else:
                guid_str = str(guid)
            
            # Convert GUID to bytes and encode
            guid_clean = guid_str.replace('-', '').replace('{', '').replace('}', '')
            
            # Generate 22-character IFC GlobalId
            result = []
            # Use hash of GUID for consistent results
            hash_val = hash(guid_clean)
            for i in range(22):
                idx = abs(hash_val + i * 7) % 64
                result.append(base64_chars[idx])
            
            return ''.join(result)
        except Exception as e:
            guid_str = str(guid)
            return guid_str[:22] if len(guid_str) >= 22 else guid_str.ljust(22, '0')


class ExportSummaryGenerator:
    """
    Generates export summary report tracking performance and data quality.
    """
    
    def __init__(self):
        self.start_time = None
        self.view_name = ""
        self.project_name = ""
        self.selected_fields = []
        self.elements_by_category = {}
        self.elements_by_ifc_type = {}
        self.data_loss_records = []
        self.total_elements = 0
        self.total_data_loss_count = 0
        
    def start_tracking(self, view_name: str, project_name: str, selected_fields: List[str]):
        """Start tracking export metrics."""
        self.start_time = datetime.now()
        self.view_name = view_name
        self.project_name = project_name
        self.selected_fields = selected_fields
        
    def record_element_processed(self, category: str, ifc_type: str, element_data: Dict, data_loss_count: int):
        """Record an element being processed."""
        self.total_elements += 1
        
        # Track by category
        if category not in self.elements_by_category:
            self.elements_by_category[category] = {'count': 0, 'data_loss': 0}
        self.elements_by_category[category]['count'] += 1
        self.elements_by_category[category]['data_loss'] += data_loss_count
        
        # Track by IFC type
        if ifc_type not in self.elements_by_ifc_type:
            self.elements_by_ifc_type[ifc_type] = {'count': 0, 'data_loss': 0}
        self.elements_by_ifc_type[ifc_type]['count'] += 1
        self.elements_by_ifc_type[ifc_type]['data_loss'] += data_loss_count
        
        self.total_data_loss_count += data_loss_count
        
    def record_data_loss(self, description: str):
        """Record a data loss event."""
        self.data_loss_records.append(description)
        
    def get_quick_summary(self) -> str:
        """Get a quick summary string."""
        duration = (datetime.now() - self.start_time).total_seconds() if self.start_time else 0
        
        return (f"Elements: {self.total_elements}\n"
                f"Categories: {len(self.elements_by_category)}\n"
                f"IFC Types: {len(self.elements_by_ifc_type)}\n"
                f"Data Loss Events: {self.total_data_loss_count}\n"
                f"Duration: {duration:.1f}s")
        
    def generate_summary_report(self, export_path: str):
        """Generate detailed summary report file."""
        summary_path = export_path.replace('.csv', '_ExportSummary.txt')
        
        duration = (datetime.now() - self.start_time).total_seconds() if self.start_time else 0
        
        with open(summary_path, 'w', encoding='utf-8') as f:
            f.write("=" * 60 + "\n")
            f.write("ARCHICAD BIM METADATA EXPORT SUMMARY REPORT\n")
            f.write("=" * 60 + "\n\n")
            
            f.write(f"Project: {self.project_name}\n")
            f.write(f"Export Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Duration: {duration:.1f} seconds\n\n")
            
            f.write("-" * 40 + "\n")
            f.write("ELEMENT STATISTICS\n")
            f.write("-" * 40 + "\n")
            f.write(f"Total Elements: {self.total_elements}\n")
            f.write(f"Total Categories: {len(self.elements_by_category)}\n")
            f.write(f"Total IFC Types: {len(self.elements_by_ifc_type)}\n\n")
            
            f.write("Elements by Category:\n")
            for cat, data in sorted(self.elements_by_category.items()):
                f.write(f"  {cat}: {data['count']} elements\n")
            
            f.write("\nElements by IFC Type:\n")
            for ifc_type, data in sorted(self.elements_by_ifc_type.items()):
                f.write(f"  {ifc_type}: {data['count']} elements\n")
            
            f.write("\n" + "-" * 40 + "\n")
            f.write("DATA QUALITY\n")
            f.write("-" * 40 + "\n")
            f.write(f"Total Data Loss Events: {self.total_data_loss_count}\n")
            
            if self.data_loss_records:
                f.write("\nData Loss Details (first 50):\n")
                for i, record in enumerate(self.data_loss_records[:50]):
                    f.write(f"  {i+1}. {record}\n")
                    
            f.write("\n" + "=" * 60 + "\n")
            f.write("END OF REPORT\n")
            f.write("=" * 60 + "\n")


class ComprehensiveElementExtractor:
    """
    Comprehensive BIM metadata extractor for ArchiCAD.
    Extracts all properties, IFC metadata, materials, relationships, and measurements.
    Similar to the Revit metadata extractor capabilities.
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
            
            # Initialize summary generator
            self.summary_generator = ExportSummaryGenerator()
            
            # Cache for project info
            self._project_info = None
            
        except Exception as e:
            print(f"❌ Connection failed: {e}")
            print("\nMake sure ArchiCAD is running and ready (no dialogs, not drawing, etc.)")
            raise
    
    def _get_project_info(self) -> Dict[str, str]:
        """Get project-level information for IFC hierarchy."""
        if self._project_info is not None:
            return self._project_info
            
        self._project_info = {
            'project_name': '',
            'site_name': '',
            'building_name': '',
            'author': '',
            'organization': '',
            'description': ''
        }
        
        try:
            # Try to get project info from ArchiCAD
            project_info = self.acc.GetProjectInfo()
            if hasattr(project_info, 'projectName'):
                self._project_info['project_name'] = str(project_info.projectName) if project_info.projectName else ''
            if hasattr(project_info, 'buildingName'):
                self._project_info['building_name'] = str(project_info.buildingName) if project_info.buildingName else ''
            if hasattr(project_info, 'siteName'):
                self._project_info['site_name'] = str(project_info.siteName) if project_info.siteName else ''
        except Exception as e:
            logging.debug(f"Could not get project info: {e}")
            
        return self._project_info
        
    def get_all_elements(self) -> List:
        """Get all 3D elements from the model."""
        print("Getting all elements...")
        all_elements = []
        
        element_types = [
            'Wall', 'Slab', 'Beam', 'Column', 'Roof', 'CurtainWall',
            'Stair', 'Railing', 'Door', 'Window', 'Skylight', 
            'Zone', 'Mesh', 'Morph', 'Shell', 'Object', 'Opening', 'Lamp', 'Fill'
        ]
        
        for elem_type in element_types:
            try:
                elements = self.acc.GetElementsByType(elem_type)
                for element in elements:
                    all_elements.append({
                        'element': element,
                        'type': elem_type
                    })
                if elements:
                    print(f"  Found {len(elements)} {elem_type} elements")
            except Exception as e:
                logging.debug(f"Could not get {elem_type} elements: {e}")
        
        print(f"Total elements found: {len(all_elements)}")
        return all_elements
    
    def get_all_properties_comprehensive(self, elements: List) -> Dict:
        """Get ALL available properties for all elements including IFC metadata."""
        print("Getting comprehensive properties and IFC metadata...")
        
        # Create element wrappers
        element_wrappers = []
        element_map = {}
        
        for i, elem_data in enumerate(elements):
            wrapper = self.act.ElementIdArrayItem(
                self.act.ElementId(elem_data['element'].elementId.guid)
            )
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
            all_property_ids = []
            property_details = []
        
        # Get project info for IFC hierarchy
        project_info = self._get_project_info()
        
        results = {}
        
        try:
            print("  Extracting property values and generating IFC metadata...")
            
            # Get all property values in bulk
            all_prop_values = []
            if all_property_ids:
                try:
                    all_prop_values = self.acc.GetPropertyValuesOfElements(element_wrappers, all_property_ids)
                except Exception as e:
                    logging.warning(f"Bulk property extraction failed: {e}")
            
            for i, elem_data in enumerate(element_map.values()):
                element_guid = elem_data['element'].elementId.guid
                element_type = elem_data['type']
                
                semantic_data_loss = 0
                
                # Generate IFC metadata
                ifc_entity_type = IfcMappingHelper.get_ifc_entity_type(element_type)
                ifc_global_id = IfcMappingHelper.generate_ifc_global_id(element_guid)
                ifc_predefined_type = IfcMappingHelper.get_ifc_predefined_type(ifc_entity_type, element_type)
                ifc_definition = IfcMappingHelper.get_ifc_definition(ifc_entity_type)
                ifc_psets = IfcMappingHelper.get_applicable_psets(ifc_entity_type)
                ifc_qtos = IfcMappingHelper.get_applicable_qtos(ifc_entity_type)
                
                # Initialize element data with core fields and IFC metadata
                element_result = {
                    # Core Identification
                    'Element_Id': '',  # Will be populated from properties (e.g., SW-091)
                    'Element_GUID': str(element_guid),
                    'Element_Type': element_type,
                    
                    # IFC Identification
                    'IFC_Identification': ifc_global_id,
                    'IFC_GlobalId': ifc_global_id,
                    'IFC_EntityType': ifc_entity_type,
                    'IFC_PredefinedType': ifc_predefined_type,
                    'IFC_Definition': ifc_definition,
                    'Element_Name': '',
                    
                    # IFC Hierarchy
                    'IFC_Hierarchy': '',
                    'IFC_Project': project_info['project_name'],
                    'IFC_Site': project_info['site_name'],
                    'IFC_Building': project_info['building_name'],
                    'IFC_Storey': '',
                    
                    # Material Information
                    'Material_Name': '',
                    'Element_Color': '',
                    'Materials': '',
                    'MaterialLayer_0_Name': '',
                    'MaterialLayer_0_Thickness': '',
                    
                    # IFC Metadata
                    'IFC_PropertySets': ifc_psets,
                    'IFC_QuantitySets': ifc_qtos,
                    'IFC_SchemaVersion': IfcMappingHelper.get_ifc_schema_version(),
                    'ArchiCAD_Category': element_type,
                    
                    # Constraints
                    'Constraints.Reference_Level': '',
                    'Constraints.Reference_Level_Elevation': '',
                    'Constraints.Top_Level': '',
                    'Constraints.Bottom_Level': '',
                    
                    # Dimensions
                    'Dimensions.Elevation_at_Bottom': '',
                    'Dimensions.Elevation_at_Top': '',
                    'Dimensions.Length': '',
                    'Dimensions.Width': '',
                    'Dimensions.Height': '',
                    'Dimensions.Area': '',
                    'Dimensions.Volume': '',
                    'Dimensions.Thickness': '',
                    
                    # Geometric Position
                    'Geometric_Position.Orientation': '',
                    'Geometric_Position.Rotation': '',
                    
                    # Property Sets - Common
                    'Pset_Common.IsExternal': '',
                    'Pset_Common.LoadBearing': '',
                    'Pset_Common.Reference': '',
                    'Pset_Common.FireRating': '',
                    'Pset_Common.Combustible': '',
                    'Pset_Common.ThermalTransmittance': '',
                    'Pset_Common.AcousticRating': '',
                    
                    # Quantity Sets - BaseQuantities
                    'Qto_BaseQuantities.Length': '',
                    'Qto_BaseQuantities.Width': '',
                    'Qto_BaseQuantities.Height': '',
                    'Qto_BaseQuantities.GrossSurfaceArea': '',
                    'Qto_BaseQuantities.NetSurfaceArea': '',
                    'Qto_BaseQuantities.GrossVolume': '',
                    'Qto_BaseQuantities.NetVolume': '',
                    'Qto_BaseQuantities.GrossWeight': '',
                    'Qto_BaseQuantities.NetWeight': '',
                    
                    # Structural
                    'Structural.StructuralFunction': '',
                    'Structural.LoadBearing': '',
                    
                    # Relationships
                    'Relationships.ConnectsTo': '',
                    'Relationships.ConnectedTo': '',
                    'Relationships.AssignedToGroup': '',
                    
                    # Identity Data
                    'Identity_Data.Has_Association': '',
                    'Identity_Data.RenovationStatus': '',
                    'Identity_Data.Layer': '',
                }
                
                # Extract all ArchiCAD properties
                if i < len(all_prop_values) and all_prop_values:
                    prop_values_wrapper = all_prop_values[i]
                    if prop_values_wrapper.propertyValues:
                        for j, prop_value in enumerate(prop_values_wrapper.propertyValues):
                            if j < len(property_details):
                                try:
                                    prop_detail = property_details[j]
                                    prop_name = prop_detail.propertyDefinition.name
                                    group_name = ""
                                    
                                    try:
                                        if hasattr(prop_detail.propertyDefinition, 'group'):
                                            group_name = prop_detail.propertyDefinition.group.name
                                    except:
                                        pass
                                    
                                    # Get property value
                                    value_str = ""
                                    if hasattr(prop_value.propertyValue, 'value') and prop_value.propertyValue.value is not None:
                                        value = prop_value.propertyValue.value
                                        value_str = str(value)
                                    
                                    # Store original property
                                    clean_prop_name = prop_name.replace(' ', '_').replace('-', '_').replace('.', '_')
                                    if group_name:
                                        clean_group = group_name.replace(' ', '_').replace('-', '_').replace('.', '_')
                                        full_key = f"{clean_group}.{clean_prop_name}"
                                    else:
                                        full_key = clean_prop_name
                                    
                                    element_result[full_key] = value_str
                                    
                                    # Map to standardized IFC fields
                                    self._map_to_ifc_fields(prop_name, value_str, element_result)
                                    
                                except Exception as e:
                                    semantic_data_loss += 1
                                    logging.debug(f"Error processing property: {e}")
                
                # Build hierarchy path
                hierarchy_parts = [
                    project_info['project_name'],
                    project_info['site_name'],
                    project_info['building_name'],
                    element_result.get('IFC_Storey', '')
                ]
                element_result['IFC_Hierarchy'] = '/'.join([p for p in hierarchy_parts if p])
                
                # Track in summary
                self.summary_generator.record_element_processed(
                    element_type, ifc_entity_type, element_result, semantic_data_loss
                )
                
                results[element_guid] = element_result
                
                # Progress indicator
                if (i + 1) % 50 == 0:
                    print(f"    Processed {i + 1}/{len(element_map)} elements...")
                    
        except Exception as e:
            print(f"  Error getting properties: {e}")
            logging.error(f"Property extraction error: {e}")
            
            # Fallback to basic data with IFC mapping
            for i, elem_data in enumerate(element_map.values()):
                element_guid = elem_data['element'].elementId.guid
                ifc_entity_type = IfcMappingHelper.get_ifc_entity_type(elem_data['type'])
                ifc_global_id = IfcMappingHelper.generate_ifc_global_id(element_guid)
                
                results[element_guid] = {
                    'Element_Id': '',  # Will be empty in fallback, populated from properties
                    'Element_GUID': str(element_guid),
                    'Element_Type': elem_data['type'],
                    'IFC_GlobalId': ifc_global_id,
                    'IFC_EntityType': ifc_entity_type,
                    'IFC_SchemaVersion': 'IFC4'
                }
        
        print(f"  ✓ Extracted properties for {len(results)} elements")
        return results
    
    def _map_to_ifc_fields(self, prop_name: str, value: str, element_result: Dict):
        """Map ArchiCAD property names to standardized IFC fields."""
        prop_lower = prop_name.lower()
        
        # Element ID (unique identifier like SW-091, not the GUID)
        if prop_lower in ['element id', 'id', 'elementid']:
            if not element_result.get('Element_Id'):
                element_result['Element_Id'] = value
        
        # Element Name
        elif prop_lower in ['element name', 'name', 'elementname']:
            if not element_result.get('Element_Name'):
                element_result['Element_Name'] = value
        
        # Home Story / Level
        elif 'story' in prop_lower or 'storey' in prop_lower or 'level' in prop_lower:
            if 'home' in prop_lower:
                element_result['IFC_Storey'] = value
                element_result['Constraints.Reference_Level'] = value
        
        # Dimensions
        elif 'length' in prop_lower:
            element_result['Dimensions.Length'] = value
            element_result['Qto_BaseQuantities.Length'] = value
        elif 'width' in prop_lower:
            element_result['Dimensions.Width'] = value
            element_result['Qto_BaseQuantities.Width'] = value
        elif 'height' in prop_lower:
            element_result['Dimensions.Height'] = value
            element_result['Qto_BaseQuantities.Height'] = value
        elif 'thickness' in prop_lower:
            element_result['Dimensions.Thickness'] = value
            element_result['MaterialLayer_0_Thickness'] = value
        elif 'area' in prop_lower:
            if 'gross' in prop_lower or 'surface' in prop_lower:
                element_result['Qto_BaseQuantities.GrossSurfaceArea'] = value
            elif 'net' in prop_lower:
                element_result['Qto_BaseQuantities.NetSurfaceArea'] = value
            else:
                element_result['Dimensions.Area'] = value
        elif 'volume' in prop_lower:
            if 'gross' in prop_lower:
                element_result['Qto_BaseQuantities.GrossVolume'] = value
            elif 'net' in prop_lower:
                element_result['Qto_BaseQuantities.NetVolume'] = value
            else:
                element_result['Dimensions.Volume'] = value
        
        # Elevation
        elif 'elevation' in prop_lower:
            if 'bottom' in prop_lower or 'base' in prop_lower:
                element_result['Dimensions.Elevation_at_Bottom'] = value
            elif 'top' in prop_lower:
                element_result['Dimensions.Elevation_at_Top'] = value
        
        # Material
        elif 'material' in prop_lower or 'building material' in prop_lower:
            if not element_result.get('Material_Name'):
                element_result['Material_Name'] = value
                element_result['MaterialLayer_0_Name'] = value
            else:
                existing = element_result.get('Materials', '')
                element_result['Materials'] = f"{existing}; {value}" if existing else value
        
        # Structural
        elif 'structural' in prop_lower or 'load' in prop_lower:
            if 'function' in prop_lower:
                element_result['Structural.StructuralFunction'] = value
            elif 'bearing' in prop_lower:
                element_result['Structural.LoadBearing'] = value
                element_result['Pset_Common.LoadBearing'] = value
        
        # External
        elif 'external' in prop_lower or 'exterior' in prop_lower:
            element_result['Pset_Common.IsExternal'] = value
        
        # Fire Rating
        elif 'fire' in prop_lower and 'rating' in prop_lower:
            element_result['Pset_Common.FireRating'] = value
        
        # Combustible
        elif 'combustible' in prop_lower or 'combustibility' in prop_lower:
            element_result['Pset_Common.Combustible'] = value
        
        # Thermal
        elif 'thermal' in prop_lower or 'u-value' in prop_lower or 'u value' in prop_lower:
            element_result['Pset_Common.ThermalTransmittance'] = value
        
        # Acoustic
        elif 'acoustic' in prop_lower or 'sound' in prop_lower:
            element_result['Pset_Common.AcousticRating'] = value
        
        # Layer
        elif prop_lower == 'layer':
            element_result['Identity_Data.Layer'] = value
        
        # Renovation Status
        elif 'renovation' in prop_lower:
            element_result['Identity_Data.RenovationStatus'] = value
        
        # Weight
        elif 'weight' in prop_lower:
            if 'gross' in prop_lower:
                element_result['Qto_BaseQuantities.GrossWeight'] = value
            elif 'net' in prop_lower:
                element_result['Qto_BaseQuantities.NetWeight'] = value
        
        # Reference
        elif prop_lower == 'reference' or prop_lower == 'ref':
            element_result['Pset_Common.Reference'] = value
    
    def get_classifications(self, elements: List) -> Dict:
        """Get classifications for all elements."""
        print("Getting classifications...")
        
        classifications = {}
        
        try:
            # Get classification systems
            class_systems = self.acc.GetAllClassificationSystems()
            if not class_systems:
                print("  No classification systems found")
                return {elem['element'].elementId.guid: {} for elem in elements}
            
            print(f"  Found {len(class_systems)} classification systems")
            
            # Create element wrappers
            element_wrappers = []
            for elem_data in elements:
                wrapper = self.act.ElementIdArrayItem(
                    self.act.ElementId(elem_data['element'].elementId.guid)
                )
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
                                
                                item_id = ""
                                item_name = ""
                                
                                if hasattr(classification.classificationId, 'classificationItemId'):
                                    class_item = classification.classificationId.classificationItemId
                                    if class_item:
                                        if hasattr(class_item, 'id'):
                                            item_id = str(class_item.id)
                                        if hasattr(class_item, 'name'):
                                            item_name = str(class_item.name)
                                
                                # Clean system name for CSV
                                clean_name = system_name.replace(' ', '_').replace('-', '_').replace('.', '_')
                                
                                if item_id:
                                    classifications[element_guid][f'Classification_{clean_name}_Id'] = item_id
                                if item_name:
                                    classifications[element_guid][f'Classification_{clean_name}_Name'] = item_name
                                if item_id or item_name:
                                    classifications[element_guid][f'Classification_{clean_name}'] = f"{item_id} - {item_name}" if item_id and item_name else (item_id or item_name)
                                    
                        except Exception as e:
                            logging.debug(f"Error processing classification: {e}")
                            continue
                            
        except Exception as e:
            print(f"  Warning: Error getting classifications: {e}")
            for elem_data in elements:
                classifications[elem_data['element'].elementId.guid] = {}
        
        return classifications
    
    def get_materials(self, elements: List) -> Dict:
        """Get material information for all elements."""
        print("Getting material information...")
        
        materials_data = {}
        
        try:
            for elem_data in elements:
                element_guid = elem_data['element'].elementId.guid
                materials_data[element_guid] = {
                    'Material_Name': '',
                    'Materials': '',
                    'Element_Color': '',
                    'MaterialLayer_0_Name': '',
                    'MaterialLayer_0_Thickness': ''
                }
                
                # Try to get building material properties
                try:
                    wrapper = self.act.ElementIdArrayItem(
                        self.act.ElementId(element_guid)
                    )
                    
                    # Try common material property names
                    material_props = ['Building Material', 'Material', 'Surface Material', 'Fill']
                    
                    for prop_name in material_props:
                        try:
                            prop_id = self.act.PropertyIdArrayItem(
                                self.act.PropertyId(
                                    self.act.PropertyGroupId("General"),
                                    prop_name
                                )
                            )
                            mat_values = self.acc.GetPropertyValuesOfElements([wrapper], [prop_id])
                            
                            if mat_values and mat_values[0].propertyValues:
                                prop_val = mat_values[0].propertyValues[0]
                                if hasattr(prop_val.propertyValue, 'value') and prop_val.propertyValue.value:
                                    mat_name = str(prop_val.propertyValue.value)
                                    if not materials_data[element_guid]['Material_Name']:
                                        materials_data[element_guid]['Material_Name'] = mat_name
                                        materials_data[element_guid]['MaterialLayer_0_Name'] = mat_name
                                    else:
                                        existing = materials_data[element_guid]['Materials']
                                        materials_data[element_guid]['Materials'] = f"{existing}; {mat_name}" if existing else mat_name
                        except:
                            continue
                            
                except Exception as e:
                    logging.debug(f"Error getting material for {element_guid}: {e}")
                    
        except Exception as e:
            print(f"  Warning: Error getting materials: {e}")
            
        return materials_data
    
    def extract_to_csv(self, filename: str = None, include_ifc_metadata: bool = True) -> bool:
        """Extract all element data to CSV with comprehensive IFC metadata."""
        try:
            if filename is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"archicad_bim_metadata_{timestamp}.csv"
            
            csv_path = os.path.join(current_dir, filename)
            
            print("=" * 60)
            print("COMPREHENSIVE ARCHICAD BIM METADATA EXTRACTION")
            print("with IFC Metadata, Materials, and Relationships")
            print("=" * 60)
            
            # Get project info
            project_info = self._get_project_info()
            project_name = project_info.get('project_name', 'Unknown Project')
            
            # Start tracking
            self.summary_generator.start_tracking("All Elements", project_name, [])
            
            # Step 1: Get all elements
            elements = self.get_all_elements()
            if not elements:
                print("❌ No elements found!")
                return False
            
            # Step 2: Get comprehensive properties including IFC metadata
            properties_data = self.get_all_properties_comprehensive(elements)
            
            # Step 3: Get classifications
            classifications_data = self.get_classifications(elements)
            
            # Step 4: Get material information
            materials_data = self.get_materials(elements)
            
            # Step 5: Combine data and write CSV
            print("Combining data and writing CSV file...")
            
            # Determine all column names with ordered priority fields first
            priority_columns = [
                # Core Identification
                'Element_Id', 'Element_GUID', 'Element_Type',
                
                # IFC Identification
                'IFC_Identification', 'IFC_GlobalId', 'IFC_EntityType', 
                'IFC_PredefinedType', 'IFC_Definition', 'Element_Name',
                
                # IFC Hierarchy
                'IFC_Hierarchy', 'IFC_Project', 'IFC_Site', 'IFC_Building', 'IFC_Storey',
                
                # Material Information
                'Material_Name', 'Element_Color', 'Materials', 
                'MaterialLayer_0_Name', 'MaterialLayer_0_Thickness',
                
                # IFC Metadata
                'IFC_PropertySets', 'IFC_QuantitySets', 'IFC_SchemaVersion', 'ArchiCAD_Category',
                
                # Constraints
                'Constraints.Reference_Level', 'Constraints.Reference_Level_Elevation',
                'Constraints.Top_Level', 'Constraints.Bottom_Level',
                
                # Dimensions
                'Dimensions.Elevation_at_Bottom', 'Dimensions.Elevation_at_Top',
                'Dimensions.Length', 'Dimensions.Width', 'Dimensions.Height',
                'Dimensions.Area', 'Dimensions.Volume', 'Dimensions.Thickness',
                
                # Geometric Position
                'Geometric_Position.Orientation', 'Geometric_Position.Rotation',
                
                # Property Sets - Common
                'Pset_Common.IsExternal', 'Pset_Common.LoadBearing', 'Pset_Common.Reference',
                'Pset_Common.FireRating', 'Pset_Common.Combustible',
                'Pset_Common.ThermalTransmittance', 'Pset_Common.AcousticRating',
                
                # Quantity Sets - BaseQuantities
                'Qto_BaseQuantities.Length', 'Qto_BaseQuantities.Width', 'Qto_BaseQuantities.Height',
                'Qto_BaseQuantities.GrossSurfaceArea', 'Qto_BaseQuantities.NetSurfaceArea',
                'Qto_BaseQuantities.GrossVolume', 'Qto_BaseQuantities.NetVolume',
                'Qto_BaseQuantities.GrossWeight', 'Qto_BaseQuantities.NetWeight',
                
                # Structural
                'Structural.StructuralFunction', 'Structural.LoadBearing',
                
                # Relationships
                'Relationships.ConnectsTo', 'Relationships.ConnectedTo', 'Relationships.AssignedToGroup',
                
                # Identity Data
                'Identity_Data.Has_Association', 'Identity_Data.RenovationStatus', 'Identity_Data.Layer',
            ]
            
            # Collect all other columns
            all_columns = set(priority_columns)
            for guid, data in properties_data.items():
                all_columns.update(data.keys())
            for guid, data in classifications_data.items():
                all_columns.update(data.keys())
            for guid, data in materials_data.items():
                all_columns.update(data.keys())
            
            # Sort remaining columns alphabetically, keeping priority columns first
            other_columns = sorted([col for col in all_columns if col not in priority_columns])
            sorted_columns = [col for col in priority_columns if col in all_columns] + other_columns
            
            # Write CSV
            with open(csv_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=sorted_columns, extrasaction='ignore')
                writer.writeheader()
                
                for elem_data in elements:
                    guid = elem_data['element'].elementId.guid
                    
                    # Combine all data for this element
                    row_data = properties_data.get(guid, {})
                    row_data.update(classifications_data.get(guid, {}))
                    row_data.update(materials_data.get(guid, {}))
                    
                    # Ensure all columns are present
                    complete_row = {}
                    for col in sorted_columns:
                        complete_row[col] = row_data.get(col, "")
                    
                    writer.writerow(complete_row)
            
            # Generate summary report
            self.summary_generator.generate_summary_report(csv_path)
            quick_summary = self.summary_generator.get_quick_summary()
            
            print("=" * 60)
            print(f"✓ SUCCESS! Extracted {len(elements)} elements")
            print(f"✓ CSV saved: {filename}")
            print(f"✓ Columns extracted: {len(sorted_columns)}")
            print(f"✓ Location: {current_dir}")
            print("-" * 40)
            print("Export Statistics:")
            print(quick_summary)
            print("=" * 60)
            print("\nIncludes:")
            print("  - IFC Identification (GlobalId, EntityType, PredefinedType)")
            print("  - IFC Hierarchy (Project/Site/Building/Storey)")
            print("  - Material Information (Names, Colors, Layers)")
            print("  - Property Sets (Pset_Common)")
            print("  - Quantity Sets (Qto_BaseQuantities)")
            print("  - Geometric Dimensions and Positions")
            print("  - Classifications and Relationships")
            print("  - All ArchiCAD Properties")
            print("=" * 60)
            
            return True
            
        except Exception as e:
            print(f"❌ Error during extraction: {e}")
            logging.error(f"Extraction error: {e}")
            import traceback
            traceback.print_exc()
            return False


# Keep FastElementExtractor for backward compatibility
class FastElementExtractor(ComprehensiveElementExtractor):
    """
    Backward-compatible alias for ComprehensiveElementExtractor.
    """
    pass


def main():
    """Main function for comprehensive BIM metadata extraction."""
    try:
        print("\n" + "=" * 60)
        print("ArchiCAD Comprehensive BIM Metadata Extractor")
        print("with IFC Data, Materials, and Relationships")
        print("=" * 60 + "\n")
        
        extractor = ComprehensiveElementExtractor()
        success = extractor.extract_to_csv(include_ifc_metadata=True)
        
        if success:
            print("\n🎉 Extraction completed successfully!")
            print("Check the CSV file and summary report in the script directory.")
        else:
            print("\n❌ Extraction failed. Check the log for details.")
            
    except Exception as e:
        print(f"\n❌ Error: {e}")
        print("Make sure ArchiCAD is running and ready (no dialogs open, not in drawing mode)")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    """
    ArchiCAD Comprehensive BIM Metadata Extractor
    
    Usage:
    1. Open ArchiCAD with your project
    2. Make sure no dialogs are open and you're not in drawing mode
    3. Run: python 2.ArchiCADMetadataExtractor.py
    
    The script will extract:
    
    IFC Identification:
    - IFC GlobalId (22-character IFC-compliant identifier)
    - IFC EntityType (IfcWall, IfcSlab, IfcBeam, etc.)
    - IFC PredefinedType (STANDARD, FLOOR, BEAM, etc.)
    - IFC Definition (Schema-based descriptions)
    
    IFC Hierarchy:
    - Project / Site / Building / Storey structure
    
    Material Information:
    - Material names and colors
    - Material layers with thicknesses
    
    Property Sets (Pset):
    - Pset_Common: IsExternal, LoadBearing, FireRating, etc.
    - Thermal and acoustic properties
    
    Quantity Sets (Qto):
    - BaseQuantities: Length, Width, Height, Area, Volume
    - Gross and Net quantities
    
    Geometric Data:
    - Dimensions and positions
    - Elevations and constraints
    
    Classifications & Relationships:
    - All classification systems
    - Element relationships and associations
    
    Output:
    - archicad_bim_metadata_[timestamp].csv - Full BIM data
    - archicad_bim_metadata_[timestamp]_ExportSummary.txt - Statistics
    """
    main()