"""
Comprehensive PowerPoint Feature Extractor
Extracts all features from PPTX files and stores in structured XML training format
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import json
import hashlib
from datetime import datetime
from collections import defaultdict
import colorsys
import re

class PPTXFeatureExtractor:
    def __init__(self, pptx_path, output_path):
        self.pptx_path = Path(pptx_path)
        self.output_path = Path(output_path)
        self.zip_file = zipfile.ZipFile(pptx_path, 'r')
        self.namespaces = {
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
            'dc': 'http://purl.org/dc/elements/1.1/',
            'dcterms': 'http://purl.org/dc/terms/',
            'extended': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
        }
        
        # Register namespaces
        for prefix, uri in self.namespaces.items():
            ET.register_namespace(prefix, uri)
    
    def extract_all_features(self):
        """Main extraction pipeline"""
        print(f"Extracting features from {self.pptx_path.name}...")
        
        training_data = ET.Element('presentation_training_data', version='1.0')
        
        # A. Document level metadata
        doc_meta = self.extract_document_metadata()
        training_data.append(doc_meta)
        
        # B. Theme and brand style
        theme_def = self.extract_theme_definition()
        training_data.append(theme_def)
        
        # C. Slide masters and layouts
        masters = self.extract_slide_masters()
        training_data.append(masters)
        
        # D. Extract all slides
        slides_elem = ET.SubElement(training_data, 'slides')
        slide_files = self.get_slide_files()
        
        for idx, slide_file in enumerate(slide_files, 1):
            print(f"  Processing slide {idx}/{len(slide_files)}...")
            slide_elem = self.extract_slide_features(slide_file, idx)
            slides_elem.append(slide_elem)
        
        # E. Global statistics
        stats = self.compute_global_statistics(training_data)
        training_data.append(stats)
        
        # Write to file
        tree = ET.ElementTree(training_data)
        ET.indent(tree, space='  ')
        tree.write(self.output_path, encoding='utf-8', xml_declaration=True)
        print(f"✓ Training data saved to {self.output_path}")
        
        return training_data
    
    def extract_document_metadata(self):
        """Extract document-level metadata"""
        doc_meta = ET.Element('document_metadata')
        
        # Presentation ID and file info
        pres_id = ET.SubElement(doc_meta, 'presentation_id', type='uuid')
        pres_id.text = self.generate_uuid(self.pptx_path.name)
        
        filename = ET.SubElement(doc_meta, 'filename')
        filename.text = self.pptx_path.name
        
        file_hash = ET.SubElement(doc_meta, 'file_hash', algorithm='sha256')
        file_hash.text = self.compute_file_hash()
        
        timestamp = ET.SubElement(doc_meta, 'extraction_timestamp')
        timestamp.text = datetime.now().isoformat()
        
        # Provenance from core.xml
        provenance = ET.SubElement(doc_meta, 'provenance')
        core_props = self.extract_core_properties()
        for key, value in core_props.items():
            elem = ET.SubElement(provenance, key)
            elem.text = str(value)
        
        # Slide dimensions from presentation.xml
        dimensions = self.extract_slide_dimensions()
        doc_meta.append(dimensions)
        
        # Custom properties
        custom_props = self.extract_custom_properties()
        if custom_props:
            doc_meta.append(custom_props)
        
        return doc_meta
    
    def extract_core_properties(self):
        """Extract core properties from docProps/core.xml"""
        props = {}
        try:
            core_xml = self.zip_file.read('docProps/core.xml')
            root = ET.fromstring(core_xml)
            
            props['author'] = self.get_text(root, './/dc:creator', '')
            props['created_date'] = self.get_text(root, './/dcterms:created', '')
            props['modified_date'] = self.get_text(root, './/dcterms:modified', '')
            props['revision'] = self.get_text(root, './/cp:revision', '1')
            props['language'] = self.get_text(root, './/dc:language', 'en-US')
        except:
            pass
        
        # Try extended properties
        try:
            app_xml = self.zip_file.read('docProps/app.xml')
            root = ET.fromstring(app_xml)
            props['company'] = self.get_text(root, './/extended:Company', '')
        except:
            pass
        
        return props
    
    def extract_slide_dimensions(self):
        """Extract slide size from presentation.xml"""
        dimensions = ET.Element('slide_dimensions', unit='emu')
        
        try:
            pres_xml = self.zip_file.read('ppt/presentation.xml')
            root = ET.fromstring(pres_xml)
            
            sld_sz = root.find('.//p:sldSz', self.namespaces)
            if sld_sz is not None:
                width = int(sld_sz.get('cx', 9144000))
                height = int(sld_sz.get('cy', 6858000))
                
                ET.SubElement(dimensions, 'width').text = str(width)
                ET.SubElement(dimensions, 'height').text = str(height)
                ET.SubElement(dimensions, 'normalized_width').text = '1.0'
                ET.SubElement(dimensions, 'normalized_height').text = '1.0'
                
                # Calculate aspect ratio
                aspect = self.calculate_aspect_ratio(width, height)
                ET.SubElement(dimensions, 'aspect_ratio').text = aspect
        except:
            # Defaults
            ET.SubElement(dimensions, 'width').text = '9144000'
            ET.SubElement(dimensions, 'height').text = '6858000'
        
        return dimensions
    
    def extract_custom_properties(self):
        """Extract custom properties if they exist"""
        custom_props = ET.Element('custom_properties')
        
        try:
            custom_xml = self.zip_file.read('docProps/custom.xml')
            root = ET.fromstring(custom_xml)
            
            for prop in root.findall('.//{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}property'):
                name = prop.get('name')
                value_elem = prop.find('.//*[@lpwstr]')
                if value_elem is not None:
                    value = value_elem.text
                    ET.SubElement(custom_props, 'property', key=name, value=value or '')
        except:
            pass
        
        return custom_props if len(custom_props) > 0 else None
    
    def extract_theme_definition(self):
        """Extract theme colors, fonts, and effects"""
        theme_def = ET.Element('theme_definition', id='theme1')
        
        try:
            theme_xml = self.zip_file.read('ppt/theme/theme1.xml')
            root = ET.fromstring(theme_xml)
            
            # Theme name
            theme_name = root.find('.//a:theme', self.namespaces)
            if theme_name is not None:
                ET.SubElement(theme_def, 'theme_name').text = theme_name.get('name', 'Default')
            
            # Color scheme
            color_palette = self.extract_color_scheme(root)
            theme_def.append(color_palette)
            
            # Font scheme
            font_scheme = self.extract_font_scheme(root)
            theme_def.append(font_scheme)
            
            # Effect styles
            effects = self.extract_effect_styles(root)
            if effects:
                theme_def.append(effects)
        except Exception as e:
            print(f"  Warning: Could not extract theme - {e}")
        
        return theme_def
    
    def extract_color_scheme(self, theme_root):
        """Extract color palette from theme"""
        color_palette = ET.Element('color_palette')
        color_scheme = ET.SubElement(color_palette, 'color_scheme', type='main')
        
        clr_scheme = theme_root.find('.//a:clrScheme', self.namespaces)
        if clr_scheme is not None:
            color_map = {
                'dk1': 'text1',
                'lt1': 'background1',
                'dk2': 'text2',
                'lt2': 'background2',
                'accent1': 'accent1',
                'accent2': 'accent2',
                'accent3': 'accent3',
                'accent4': 'accent4',
                'accent5': 'accent5',
                'accent6': 'accent6',
                'hlink': 'hyperlink',
                'folHlink': 'followed_hyperlink'
            }
            
            for name, role in color_map.items():
                color_elem = clr_scheme.find(f'.//a:{name}', self.namespaces)
                if color_elem is not None:
                    hex_color = self.extract_color_value(color_elem)
                    if hex_color:
                        rgb = self.hex_to_rgb(hex_color)
                        lab = self.rgb_to_lab(rgb)
                        
                        color_node = ET.SubElement(color_scheme, 'color', 
                                                   name=name, 
                                                   role=role,
                                                   hex=hex_color,
                                                   rgb=f"{rgb[0]},{rgb[1]},{rgb[2]}",
                                                   lab=f"{lab[0]:.1f},{lab[1]:.1f},{lab[2]:.1f}")
        
        return color_palette
    
    def extract_font_scheme(self, theme_root):
        """Extract font definitions"""
        font_scheme = ET.Element('font_scheme')
        
        font_scheme_elem = theme_root.find('.//a:fontScheme', self.namespaces)
        if font_scheme_elem is not None:
            # Major font (headings)
            major_font = font_scheme_elem.find('.//a:majorFont/a:latin', self.namespaces)
            if major_font is not None:
                ET.SubElement(font_scheme, 'major_font', 
                            family=major_font.get('typeface', 'Calibri Light'))
            
            # Minor font (body)
            minor_font = font_scheme_elem.find('.//a:minorFont/a:latin', self.namespaces)
            if minor_font is not None:
                ET.SubElement(font_scheme, 'minor_font',
                            family=minor_font.get('typeface', 'Calibri'))
        
        return font_scheme
    
    def extract_effect_styles(self, theme_root):
        """Extract effect styles (shadows, glows, etc.)"""
        effects = ET.Element('effect_styles')
        
        # This is simplified - full extraction would parse effectStyle elements
        effect_list = theme_root.find('.//a:effectStyleLst', self.namespaces)
        if effect_list is not None:
            for idx, effect_style in enumerate(effect_list.findall('.//a:effectStyle', self.namespaces), 1):
                style_elem = ET.SubElement(effects, 'effect_style', id=f'style{idx}')
                # Add basic info
                style_elem.set('has_effects', 'true')
        
        return effects if len(effects) > 0 else None
    
    def extract_slide_masters(self):
        """Extract slide master and layout information"""
        masters_elem = ET.Element('slide_masters')
        
        try:
            # Read presentation.xml to get master references
            pres_xml = self.zip_file.read('ppt/presentation.xml')
            pres_root = ET.fromstring(pres_xml)
            
            # Get slide master list
            sld_master_id_list = pres_root.find('.//p:sldMasterIdLst', self.namespaces)
            if sld_master_id_list is not None:
                for master_id_elem in sld_master_id_list.findall('.//p:sldMasterId', self.namespaces):
                    rid = master_id_elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    
                    # Get master file from relationships
                    master_file = self.get_relationship_target('ppt/_rels/presentation.xml.rels', rid)
                    if master_file:
                        master_elem = self.extract_master_info(f'ppt/{master_file}')
                        masters_elem.append(master_elem)
        except Exception as e:
            print(f"  Warning: Could not extract masters - {e}")
        
        return masters_elem
    
    def extract_master_info(self, master_file):
        """Extract information from a single slide master"""
        master_elem = ET.Element('slide_master', id=Path(master_file).stem)
        
        try:
            master_xml = self.zip_file.read(master_file)
            master_root = ET.fromstring(master_xml)
            
            ET.SubElement(master_elem, 'master_name').text = 'Office Theme'
            
            # Get layouts
            layouts_elem = ET.SubElement(master_elem, 'layouts')
            master_rels = master_file.replace('.xml', '.xml.rels')
            
            # Find layout files from relationships
            layout_rels = self.get_all_relationships(master_rels, 'slideLayout')
            
            for layout_rel in layout_rels:
                layout_file = f"ppt/slideLayouts/{Path(layout_rel).name}"
                layout_elem = self.extract_layout_info(layout_file)
                layouts_elem.append(layout_elem)
        except:
            pass
        
        return master_elem
    
    def extract_layout_info(self, layout_file):
        """Extract layout structure and placeholders"""
        layout_elem = ET.Element('layout', id=Path(layout_file).stem)
        
        try:
            layout_xml = self.zip_file.read(layout_file)
            layout_root = ET.fromstring(layout_xml)
            
            # Layout name
            cSld = layout_root.find('.//p:cSld', self.namespaces)
            if cSld is not None:
                layout_name = cSld.get('name', 'Unknown')
                layout_elem.set('name', layout_name)
            
            # Find placeholders - FIXED: Use simpler XPath
            all_shapes = layout_root.findall('.//p:sp', self.namespaces)
            for sp in all_shapes:
                # Check if this shape has a placeholder
                ph = sp.find('.//p:ph', self.namespaces)
                if ph is not None:
                    ph_elem = self.extract_placeholder_info(sp)
                    if ph_elem is not None:
                        layout_elem.append(ph_elem)
        except Exception as e:
            print(f"  Warning: Could not extract layout {layout_file} - {e}")
        
        return layout_elem
    
    def extract_placeholder_info(self, shape_elem):
        """Extract placeholder definition from a shape"""
        ph_elem = shape_elem.find('.//p:ph', self.namespaces)
        if ph_elem is None:
            return None
        
        ph_type = ph_elem.get('type', 'body')
        ph_idx = ph_elem.get('idx', '0')
        
        placeholder = ET.Element('placeholder', type=ph_type, idx=ph_idx)
        
        # Get geometry
        xfrm = shape_elem.find('.//a:xfrm', self.namespaces)
        if xfrm is not None:
            geom = self.extract_geometry(xfrm)
            placeholder.append(geom)
        
        return placeholder
    
    def get_slide_files(self):
        """Get list of slide files in order"""
        slide_files = []
        
        try:
            pres_xml = self.zip_file.read('ppt/presentation.xml')
            root = ET.fromstring(pres_xml)
            
            sld_id_list = root.find('.//p:sldIdLst', self.namespaces)
            if sld_id_list is not None:
                for sld_id in sld_id_list.findall('.//p:sldId', self.namespaces):
                    rid = sld_id.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    
                    # Get slide file from relationships
                    slide_file = self.get_relationship_target('ppt/_rels/presentation.xml.rels', rid)
                    if slide_file:
                        slide_files.append(f'ppt/{slide_file}')
        except:
            # Fallback: find all slide files
            slide_files = [f for f in self.zip_file.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')]
            slide_files.sort()
        
        return slide_files
    
    def extract_slide_features(self, slide_file, index):
        """Extract comprehensive features from a single slide"""
        slide_elem = ET.Element('slide', id=Path(slide_file).stem, index=str(index))
        
        try:
            slide_xml = self.zip_file.read(slide_file)
            slide_root = ET.fromstring(slide_xml)
            
            # Slide metadata
            metadata = self.extract_slide_metadata(slide_root, slide_file)
            slide_elem.append(metadata)
            
            # Background
            bg = self.extract_background(slide_root)
            if bg is not None:
                slide_elem.append(bg)
            
            # Semantic analysis
            semantic_role = self.infer_slide_role(slide_root, index)
            ET.SubElement(slide_elem, 'semantic_role').text = semantic_role
            
            # Slide hash for deduplication
            slide_hash = self.compute_slide_hash(slide_xml)
            ET.SubElement(slide_elem, 'slide_hash').text = slide_hash
            
            # Extract all elements
            elements_elem = ET.SubElement(slide_elem, 'elements')
            self.extract_slide_elements(slide_root, elements_elem, slide_file)
            
            # Computed features
            computed = self.compute_slide_features(slide_elem)
            slide_elem.append(computed)
            
        except Exception as e:
            print(f"    Error extracting slide {index}: {e}")
            import traceback
            traceback.print_exc()
        
        return slide_elem
    
    def extract_slide_metadata(self, slide_root, slide_file):
        """Extract slide-level metadata"""
        metadata = ET.Element('slide_metadata')
        
        # Check if hidden
        show = slide_root.get('show', '1')
        ET.SubElement(metadata, 'is_hidden').text = 'true' if show == '0' else 'false'
        
        # Check for notes
        rels_file = slide_file.replace('.xml', '.xml.rels')
        has_notes = self.has_relationship_type(rels_file, 'notesSlide')
        ET.SubElement(metadata, 'has_notes').text = 'true' if has_notes else 'false'
        
        # Transition
        transition = slide_root.find('.//p:transition', self.namespaces)
        trans_elem = ET.SubElement(metadata, 'transition')
        if transition is not None:
            trans_elem.set('type', transition.get('spd', 'medium'))
            trans_elem.set('duration', transition.get('dur', '500'))
        else:
            trans_elem.set('type', 'none')
        
        return metadata
    
    def extract_background(self, slide_root):
        """Extract slide background"""
        bg_elem = slide_root.find('.//p:bg', self.namespaces)
        if bg_elem is None:
            return ET.Element('background', type='none')
        
        background = ET.Element('background')
        
        # Check for solid fill
        solid_fill = bg_elem.find('.//a:solidFill', self.namespaces)
        if solid_fill is not None:
            color = self.extract_color_value(solid_fill)
            if color:
                background.set('type', 'solid')
                background.set('color', color)
                return background
        
        # Check for gradient
        grad_fill = bg_elem.find('.//a:gradFill', self.namespaces)
        if grad_fill is not None:
            background.set('type', 'gradient')
            return background
        
        # Check for image
        blip_fill = bg_elem.find('.//a:blipFill', self.namespaces)
        if blip_fill is not None:
            background.set('type', 'image')
            return background
        
        background.set('type', 'none')
        return background
    
    def extract_slide_elements(self, slide_root, elements_elem, slide_file):
        """Extract all elements (shapes, images, charts, tables) from slide"""
        sp_tree = slide_root.find('.//p:spTree', self.namespaces)
        if sp_tree is None:
            return
        
        z_order = 1
        
        # Extract shapes (text boxes)
        for sp in sp_tree.findall('.//p:sp', self.namespaces):
            elem = self.extract_shape_element(sp, z_order, slide_file)
            if elem is not None:
                elements_elem.append(elem)
                z_order += 1
        
        # Extract pictures
        for pic in sp_tree.findall('.//p:pic', self.namespaces):
            elem = self.extract_picture_element(pic, z_order, slide_file)
            if elem is not None:
                elements_elem.append(elem)
                z_order += 1
        
        # Extract graphic frames (charts, tables, SmartArt)
        for gf in sp_tree.findall('.//p:graphicFrame', self.namespaces):
            elem = self.extract_graphic_frame(gf, z_order, slide_file)
            if elem is not None:
                elements_elem.append(elem)
                z_order += 1
        
        # Extract groups
        for grp in sp_tree.findall('.//p:grpSp', self.namespaces):
            group_elem = self.extract_group_element(grp, z_order, slide_file)
            if group_elem is not None:
                elements_elem.append(group_elem)
                z_order += 1
    
    def extract_shape_element(self, shape, z_order, slide_file):
        """Extract shape/textbox element"""
        nv_sp_pr = shape.find('.//p:nvSpPr', self.namespaces)
        if nv_sp_pr is None:
            return None
        
        c_nv_pr = nv_sp_pr.find('.//p:cNvPr', self.namespaces)
        shape_id = c_nv_pr.get('id', str(z_order)) if c_nv_pr is not None else str(z_order)
        shape_name = c_nv_pr.get('name', f'shape{z_order}') if c_nv_pr is not None else f'shape{z_order}'
        
        element = ET.Element('element', id=shape_id, type='textbox', z_order=str(z_order))
        
        # Geometry
        xfrm = shape.find('.//a:xfrm', self.namespaces)
        if xfrm is not None:
            geom = self.extract_geometry(xfrm)
            element.append(geom)
        
        # Placeholder info
        ph = nv_sp_pr.find('.//p:ph', self.namespaces)
        if ph is not None:
            ph_type = ph.get('type', 'body')
            ph_idx = ph.get('idx', '0')
            ET.SubElement(element, 'placeholder', type=ph_type, idx=ph_idx)
        
        # Fill
        fill = self.extract_fill_properties(shape)
        element.append(fill)
        
        # Stroke
        stroke = self.extract_stroke_properties(shape)
        element.append(stroke)
        
        # Text content
        tx_body = shape.find('.//p:txBody', self.namespaces)
        if tx_body is not None:
            text_body = self.extract_text_body(tx_body)
            element.append(text_body)
        
        # Alt text
        alt_text = c_nv_pr.get('descr', '') if c_nv_pr is not None else ''
        if alt_text:
            acc = ET.SubElement(element, 'accessibility')
            ET.SubElement(acc, 'alt_text').text = alt_text
        
        return element
    
    def extract_picture_element(self, pic, z_order, slide_file):
        """Extract image/picture element"""
        nv_pic_pr = pic.find('.//p:nvPicPr', self.namespaces)
        if nv_pic_pr is None:
            return None
        
        c_nv_pr = nv_pic_pr.find('.//p:cNvPr', self.namespaces)
        pic_id = c_nv_pr.get('id', str(z_order)) if c_nv_pr is not None else str(z_order)
        
        element = ET.Element('element', id=pic_id, type='picture', z_order=str(z_order))
        
        # Geometry
        xfrm = pic.find('.//a:xfrm', self.namespaces)
        if xfrm is not None:
            geom = self.extract_geometry(xfrm)
            element.append(geom)
        
        # Image reference
        blip = pic.find('.//a:blip', self.namespaces)
        if blip is not None:
            embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if embed:
                rels_file = slide_file.replace('.xml', '.xml.rels')
                image_file = self.get_relationship_target(rels_file, embed)
                if image_file:
                    media = ET.SubElement(element, 'media_reference')
                    media.set('file', image_file)
                    media.set('type', Path(image_file).suffix[1:])
        
        # Alt text
        alt_text = c_nv_pr.get('descr', '') if c_nv_pr is not None else ''
        if alt_text:
            acc = ET.SubElement(element, 'accessibility')
            ET.SubElement(acc, 'alt_text').text = alt_text
        
        return element
    
    def extract_graphic_frame(self, gf, z_order, slide_file):
        """Extract chart, table, or SmartArt"""
        # Determine type
        chart = gf.find('.//c:chart', {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'})
        table = gf.find('.//a:tbl', self.namespaces)
        
        if chart is not None:
            return self.extract_chart_element(gf, z_order, slide_file)
        elif table is not None:
            return self.extract_table_element(gf, z_order, slide_file)
        else:
            # SmartArt or other
            element = ET.Element('element', type='graphic', z_order=str(z_order))
            return element
    
    def extract_chart_element(self, gf, z_order, slide_file):
        """Extract chart element"""
        element = ET.Element('element', type='chart', z_order=str(z_order))
        
        # Geometry
        xfrm = gf.find('.//a:xfrm', self.namespaces)
        if xfrm is not None:
            geom = self.extract_geometry(xfrm)
            element.append(geom)
        
        # Chart reference
        chart_elem = gf.find('.//c:chart', {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'})
        if chart_elem is not None:
            rid = chart_elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            if rid:
                rels_file = slide_file.replace('.xml', '.xml.rels')
                chart_file = self.get_relationship_target(rels_file, rid)
                if chart_file:
                    chart_def = self.extract_chart_definition(f'ppt/{chart_file}')
                    element.append(chart_def)
        
        return element
    
    def extract_chart_definition(self, chart_file):
        """Extract chart definition and data"""
        chart_def = ET.Element('chart_definition')
        
        try:
            chart_xml = self.zip_file.read(chart_file)
            chart_root = ET.fromstring(chart_xml)
            
            # Chart type
            plot_area = chart_root.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}plotArea')
            if plot_area is not None:
                # Find chart type (barChart, lineChart, pieChart, etc.)
                for child in plot_area:
                    if 'Chart' in child.tag:
                        chart_type = child.tag.split('}')[1]
                        ET.SubElement(chart_def, 'chart_type').text = chart_type
                        break
            
            # Title
            title = chart_root.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}title')
            if title is not None:
                title_text = self.get_text_from_chart_element(title)
                if title_text:
                    ET.SubElement(chart_def, 'title').text = title_text
            
            # Legend
            legend = chart_root.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}legend')
            if legend is not None:
                leg_pos = legend.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}legendPos')
                position = leg_pos.get('val', 'r') if leg_pos is not None else 'r'
                ET.SubElement(chart_def, 'legend', position=position, show='true')
            
            # Series data (simplified extraction)
            series_list = chart_root.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}ser')
            if series_list:
                series_elem = ET.SubElement(chart_def, 'series')
                for idx, ser in enumerate(series_list):
                    ser_name = self.get_text_from_chart_element(ser.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}tx'))
                    data_series = ET.SubElement(series_elem, 'data_series', idx=str(idx))
                    if ser_name:
                        data_series.set('name', ser_name)
        
        except Exception as e:
            print(f"    Warning: Could not fully extract chart - {e}")
        
        return chart_def
    
    def extract_table_element(self, gf, z_order, slide_file):
        """Extract table element"""
        element = ET.Element('element', type='table', z_order=str(z_order))
        
        # Geometry
        xfrm = gf.find('.//a:xfrm', self.namespaces)
        if xfrm is not None:
            geom = self.extract_geometry(xfrm)
            element.append(geom)
        
        # Table structure
        tbl = gf.find('.//a:tbl', self.namespaces)
        if tbl is not None:
            table_def = ET.SubElement(element, 'table_definition')
            
            # Row count
            rows = tbl.findall('.//a:tr', self.namespaces)
            ET.SubElement(table_def, 'row_count').text = str(len(rows))
            
            # Column count (from first row)
            if rows:
                cells = rows[0].findall('.//a:tc', self.namespaces)
                ET.SubElement(table_def, 'col_count').text = str(len(cells))
            
            # Extract cell content (simplified)
            rows_elem = ET.SubElement(table_def, 'rows')
            for row_idx, row in enumerate(rows):
                row_elem = ET.SubElement(rows_elem, 'row', idx=str(row_idx))
                cells = row.findall('.//a:tc', self.namespaces)
                for cell_idx, cell in enumerate(cells):
                    cell_elem = ET.SubElement(row_elem, 'cell', idx=str(cell_idx))
                    # Get cell text
                    text = self.get_all_text_from_element(cell)
                    if text:
                        ET.SubElement(cell_elem, 'text').text = text
        
        return element
    
    def extract_group_element(self, grp, z_order, slide_file):
        """Extract grouped shapes"""
        element = ET.Element('element', type='group', z_order=str(z_order))
        
        # Geometry
        xfrm = grp.find('.//a:xfrm', self.namespaces)
        if xfrm is not None:
            geom = self.extract_geometry(xfrm)
            element.append(geom)
        
        # Extract child elements
        children_elem = ET.SubElement(element, 'children')
        child_z = 1
        
        for sp in grp.findall('.//p:sp', self.namespaces):
            child = self.extract_shape_element(sp, child_z, slide_file)
            if child is not None:
                children_elem.append(child)
                child_z += 1
        
        return element
    
    def extract_geometry(self, xfrm):
        """Extract position and size from transform element"""
        geom = ET.Element('geometry', unit='normalized')
        
        # Get slide dimensions for normalization
        slide_width = 9144000  # default
        slide_height = 6858000
        
        try:
            pres_xml = self.zip_file.read('ppt/presentation.xml')
            root = ET.fromstring(pres_xml)
            sld_sz = root.find('.//p:sldSz', self.namespaces)
            if sld_sz is not None:
                slide_width = int(sld_sz.get('cx', 9144000))
                slide_height = int(sld_sz.get('cy', 6858000))
        except:
            pass
        
        # Position
        off = xfrm.find('.//a:off', self.namespaces)
        if off is not None:
            x = int(off.get('x', 0))
            y = int(off.get('y', 0))
            ET.SubElement(geom, 'x').text = f"{x / slide_width:.6f}"
            ET.SubElement(geom, 'y').text = f"{y / slide_height:.6f}"
        
        # Size
        ext = xfrm.find('.//a:ext', self.namespaces)
        if ext is not None:
            cx = int(ext.get('cx', 0))
            cy = int(ext.get('cy', 0))
            ET.SubElement(geom, 'width').text = f"{cx / slide_width:.6f}"
            ET.SubElement(geom, 'height').text = f"{cy / slide_height:.6f}"
        
        # Rotation
        rot = xfrm.get('rot', '0')
        ET.SubElement(geom, 'rotation').text = str(int(rot) / 60000)  # Convert from 1/60000 degrees
        
        # Flip
        flip_h = xfrm.get('flipH', '0')
        flip_v = xfrm.get('flipV', '0')
        ET.SubElement(geom, 'flip_h').text = 'true' if flip_h == '1' else 'false'
        ET.SubElement(geom, 'flip_v').text = 'true' if flip_v == '1' else 'false'
        
        return geom
    
    def extract_fill_properties(self, shape):
        """Extract fill properties"""
        fill = ET.Element('fill')
        
        sp_pr = shape.find('.//p:spPr', self.namespaces)
        if sp_pr is None:
            fill.set('type', 'none')
            return fill
        
        # Solid fill
        solid_fill = sp_pr.find('.//a:solidFill', self.namespaces)
        if solid_fill is not None:
            color = self.extract_color_value(solid_fill)
            if color:
                fill.set('type', 'solid')
                fill.set('color', color)
                return fill
        
        # Gradient fill
        grad_fill = sp_pr.find('.//a:gradFill', self.namespaces)
        if grad_fill is not None:
            fill.set('type', 'gradient')
            # Extract gradient stops
            stops = grad_fill.findall('.//a:gs', self.namespaces)
            for stop in stops:
                pos = stop.get('pos', '0')
                color = self.extract_color_value(stop)
                if color:
                    stop_elem = ET.SubElement(fill, 'gradient_stop', position=pos, color=color)
            return fill
        
        # Pattern fill
        patt_fill = sp_pr.find('.//a:pattFill', self.namespaces)
        if patt_fill is not None:
            fill.set('type', 'pattern')
            return fill
        
        # Image fill
        blip_fill = sp_pr.find('.//a:blipFill', self.namespaces)
        if blip_fill is not None:
            fill.set('type', 'image')
            return fill
        
        # No fill
        no_fill = sp_pr.find('.//a:noFill', self.namespaces)
        if no_fill is not None:
            fill.set('type', 'none')
            return fill
        
        fill.set('type', 'default')
        return fill
    
    def extract_stroke_properties(self, shape):
        """Extract stroke/border properties"""
        stroke = ET.Element('stroke')
        
        sp_pr = shape.find('.//p:spPr', self.namespaces)
        if sp_pr is None:
            stroke.set('width', '0')
            return stroke
        
        ln = sp_pr.find('.//a:ln', self.namespaces)
        if ln is not None:
            width = ln.get('w', '0')
            stroke.set('width', str(int(width) / 12700))  # Convert EMU to points
            
            # Stroke color
            solid_fill = ln.find('.//a:solidFill', self.namespaces)
            if solid_fill is not None:
                color = self.extract_color_value(solid_fill)
                if color:
                    stroke.set('color', color)
            
            # Dash style
            dash = ln.find('.//a:prstDash', self.namespaces)
            if dash is not None:
                stroke.set('dash', dash.get('val', 'solid'))
        else:
            stroke.set('width', '0')
        
        return stroke
    
    def extract_text_body(self, tx_body):
        """Extract text content with formatting"""
        text_body = ET.Element('text_body')
        
        # Language
        body_pr = tx_body.find('.//a:bodyPr', self.namespaces)
        if body_pr is not None:
            lang = body_pr.get('lang', 'en-US')
            ET.SubElement(text_body, 'language').text = lang
        
        # Paragraphs
        paragraphs = tx_body.findall('.//a:p', self.namespaces)
        for p_idx, p in enumerate(paragraphs):
            para_elem = self.extract_paragraph(p, p_idx)
            text_body.append(para_elem)
        
        return text_body
    
    def extract_paragraph(self, p, idx):
        """Extract paragraph with runs"""
        para = ET.Element('paragraph', idx=str(idx))
        
        # Paragraph properties
        p_pr = p.find('.//a:pPr', self.namespaces)
        if p_pr is not None:
            # Alignment
            algn = p_pr.get('algn', 'left')
            ET.SubElement(para, 'alignment').text = algn
            
            # Level
            lvl = p_pr.get('lvl', '0')
            para.set('level', lvl)
            
            # Line spacing
            ln_spc = p_pr.find('.//a:lnSpc', self.namespaces)
            if ln_spc is not None:
                spc_pct = ln_spc.find('.//a:spcPct', self.namespaces)
                if spc_pct is not None:
                    val = spc_pct.get('val', '100000')
                    ET.SubElement(para, 'line_spacing').text = str(int(val) / 100000)
            
            # Spacing before/after
            spc_bef = p_pr.find('.//a:spcBef', self.namespaces)
            if spc_bef is not None:
                spc_pts = spc_bef.find('.//a:spcPts', self.namespaces)
                if spc_pts is not None:
                    ET.SubElement(para, 'space_before').text = spc_pts.get('val', '0')
            
            spc_aft = p_pr.find('.//a:spcAft', self.namespaces)
            if spc_aft is not None:
                spc_pts = spc_aft.find('.//a:spcPts', self.namespaces)
                if spc_pts is not None:
                    ET.SubElement(para, 'space_after').text = spc_pts.get('val', '0')
            
            # Bullet/numbering
            bu_none = p_pr.find('.//a:buNone', self.namespaces)
            if bu_none is None:
                bu_char = p_pr.find('.//a:buChar', self.namespaces)
                if bu_char is not None:
                    para.set('bullet', bu_char.get('char', '•'))
                else:
                    para.set('bullet', '•')
        
        # Text runs
        runs = p.findall('.//a:r', self.namespaces)
        for r_idx, r in enumerate(runs):
            run_elem = self.extract_text_run(r, r_idx)
            para.append(run_elem)
        
        # Handle end paragraph run (for empty paragraphs)
        if not runs:
            end_para = p.find('.//a:endParaRPr', self.namespaces)
            if end_para is not None:
                run_elem = ET.Element('text_run', idx='0')
                ET.SubElement(run_elem, 'text').text = ''
                para.append(run_elem)
        
        return para
    
    def extract_text_run(self, r, idx):
        """Extract text run with character formatting"""
        run = ET.Element('text_run', idx=str(idx))
        
        # Text content
        t = r.find('.//a:t', self.namespaces)
        text = t.text if t is not None and t.text else ''
        ET.SubElement(run, 'text').text = text
        
        # Run properties
        r_pr = r.find('.//a:rPr', self.namespaces)
        if r_pr is not None:
            # Font
            latin = r_pr.find('.//a:latin', self.namespaces)
            font_family = latin.get('typeface', 'Calibri') if latin is not None else 'Calibri'
            
            # Font size (in points)
            sz = r_pr.get('sz', '1800')
            font_size = int(sz) / 100
            
            # Bold, italic, underline
            bold = r_pr.get('b', '0') == '1'
            italic = r_pr.get('i', '0') == '1'
            underline = r_pr.get('u', 'none') != 'none'
            strike = r_pr.get('strike', 'noStrike') != 'noStrike'
            
            font = ET.SubElement(run, 'font',
                                family=font_family,
                                size=str(font_size),
                                bold=str(bold).lower(),
                                italic=str(italic).lower(),
                                underline=str(underline).lower(),
                                strike=str(strike).lower())
            
            # Color
            solid_fill = r_pr.find('.//a:solidFill', self.namespaces)
            if solid_fill is not None:
                color = self.extract_color_value(solid_fill)
                if color:
                    rgb = self.hex_to_rgb(color)
                    lab = self.rgb_to_lab(rgb)
                    color_elem = ET.SubElement(run, 'color',
                                              hex=color,
                                              rgb=f"{rgb[0]},{rgb[1]},{rgb[2]}",
                                              lab=f"{lab[0]:.1f},{lab[1]:.1f},{lab[2]:.1f}")
            
            # Letter spacing
            spc = r_pr.get('spc', '0')
            if spc != '0':
                ET.SubElement(run, 'letter_spacing').text = str(int(spc) / 100)
            
            # Baseline offset (superscript/subscript)
            baseline = r_pr.get('baseline', '0')
            if baseline != '0':
                ET.SubElement(run, 'baseline_offset').text = baseline
        
        return run
    
    def extract_color_value(self, element):
        """Extract color from various color element types"""
        # sRGB color
        srgb_clr = element.find('.//a:srgbClr', self.namespaces)
        if srgb_clr is not None:
            return '#' + srgb_clr.get('val', 'FFFFFF').upper()
        
        # Scheme color
        scheme_clr = element.find('.//a:schemeClr', self.namespaces)
        if scheme_clr is not None:
            # Return scheme reference (would need theme mapping)
            val = scheme_clr.get('val', 'dk1')
            return f'scheme:{val}'
        
        # System color
        sys_clr = element.find('.//a:sysClr', self.namespaces)
        if sys_clr is not None:
            last_clr = sys_clr.get('lastClr')
            if last_clr:
                return '#' + last_clr.upper()
        
        # Preset color
        prst_clr = element.find('.//a:prstClr', self.namespaces)
        if prst_clr is not None:
            return f'preset:{prst_clr.get("val", "black")}'
        
        return None
    
    def infer_slide_role(self, slide_root, index):
        """Infer semantic role of slide"""
        # First slide is usually title
        if index == 1:
            return 'title_slide'
        
        # Count elements
        sp_tree = slide_root.find('.//p:spTree', self.namespaces)
        if sp_tree is None:
            return 'content'
        
        # Check for charts
        charts = sp_tree.findall('.//c:chart', {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'})
        if charts:
            return 'data_visualization'
        
        # Check for tables
        tables = sp_tree.findall('.//a:tbl', self.namespaces)
        if tables:
            return 'table_content'
        
        # Check for images
        pics = sp_tree.findall('.//p:pic', self.namespaces)
        if len(pics) > 2:
            return 'image_gallery'
        
        # Default
        return 'content'
    
    def compute_slide_features(self, slide_elem):
        """Compute derived features for ML training"""
        computed = ET.Element('computed_features')
        
        elements = slide_elem.find('.//elements')
        if elements is None:
            return computed
        
        element_count = len(elements)
        ET.SubElement(computed, 'element_count').text = str(element_count)
        
        # Count text vs images
        text_count = len(elements.findall('.//element[@type="textbox"]'))
        image_count = len(elements.findall('.//element[@type="picture"]'))
        chart_count = len(elements.findall('.//element[@type="chart"]'))
        
        # Text to image ratio
        if image_count + chart_count > 0:
            ratio = text_count / (image_count + chart_count)
        else:
            ratio = float('inf') if text_count > 0 else 0
        ET.SubElement(computed, 'text_to_image_ratio').text = f"{ratio:.2f}"
        
        # Whitespace ratio (simplified - based on element coverage)
        total_area = 0
        for elem in elements.findall('.//element'):
            geom = elem.find('.//geometry')
            if geom is not None:
                width_elem = geom.find('.//width')
                height_elem = geom.find('.//height')
                if width_elem is not None and height_elem is not None:
                    width = float(width_elem.text or 0)
                    height = float(height_elem.text or 0)
                    total_area += width * height
        
        whitespace = max(0, 1.0 - min(total_area, 1.0))
        ET.SubElement(computed, 'whitespace_ratio').text = f"{whitespace:.2f}"
        
        # Readability (simplified - character count)
        total_chars = 0
        for text_elem in slide_elem.findall('.//text'):
            if text_elem.text:
                total_chars += len(text_elem.text)
        
        # Flesch reading ease (simplified estimate)
        readability = max(0, min(100, 100 - (total_chars / 10)))
        ET.SubElement(computed, 'readability_flesch').text = f"{readability:.1f}"
        
        # Visual hierarchy score (title prominence)
        hierarchy_score = 0.5  # default
        title_elements = list(elements.findall('.//element'))
        for elem in title_elements:
            ph = elem.find('.//placeholder[@type="title"]')
            if ph is None:
                ph = elem.find('.//placeholder[@type="ctrTitle"]')
            if ph is not None:
                hierarchy_score = 0.9
                break
        
        ET.SubElement(computed, 'visual_hierarchy_score').text = f"{hierarchy_score:.2f}"
        
        # Color diversity (count unique colors)
        colors = set()
        for color_elem in slide_elem.findall('.//color[@hex]'):
            hex_val = color_elem.get('hex')
            if hex_val and not hex_val.startswith('scheme:'):
                colors.add(hex_val)
        color_diversity = min(len(colors) / 10, 1.0)
        ET.SubElement(computed, 'color_diversity').text = f"{color_diversity:.2f}"
        
        # Layout balance (simplified - check distribution)
        left_weight = 0
        right_weight = 0
        top_weight = 0
        bottom_weight = 0
        
        for elem in elements.findall('.//element'):
            geom = elem.find('.//geometry')
            if geom is not None:
                x_elem = geom.find('.//x')
                y_elem = geom.find('.//y')
                width_elem = geom.find('.//width')
                height_elem = geom.find('.//height')
                
                if all([x_elem is not None, y_elem is not None, width_elem is not None, height_elem is not None]):
                    x = float(x_elem.text or 0)
                    y = float(y_elem.text or 0)
                    width = float(width_elem.text or 0)
                    height = float(height_elem.text or 0)
                    
                    weight = width * height
                    
                    if x + width/2 < 0.5:
                        left_weight += weight
                    else:
                        right_weight += weight
                    
                    if y + height/2 < 0.5:
                        top_weight += weight
                    else:
                        bottom_weight += weight
        
        total = left_weight + right_weight
        h_balance = 1.0 - abs(left_weight - right_weight) / total if total > 0 else 1.0
        
        total = top_weight + bottom_weight
        v_balance = 1.0 - abs(top_weight - bottom_weight) / total if total > 0 else 1.0
        
        ET.SubElement(computed, 'layout_balance_horizontal').text = f"{h_balance:.2f}"
        ET.SubElement(computed, 'layout_balance_vertical').text = f"{v_balance:.2f}"
        
        return computed
    
    def compute_global_statistics(self, training_data):
        """Compute presentation-level statistics"""
        stats = ET.Element('global_statistics')
        
        slides = training_data.find('.//slides')
        if slides is None:
            return stats
        
        slide_count = len(slides.findall('.//slide'))
        ET.SubElement(stats, 'total_slides').text = str(slide_count)
        
        # Layout usage
        layout_usage = defaultdict(int)
        for slide in slides.findall('.//slide'):
            layout_ref = slide.get('layout_ref', 'unknown')
            layout_usage[layout_ref] += 1
        
        layout_elem = ET.SubElement(stats, 'layout_usage')
        for layout, count in layout_usage.items():
            ET.SubElement(layout_elem, 'layout', ref=layout, count=str(count))
        
        # Semantic roles
        role_usage = defaultdict(int)
        for role in slides.findall('.//semantic_role'):
            if role.text:
                role_usage[role.text] += 1
        
        role_elem = ET.SubElement(stats, 'semantic_role_distribution')
        for role, count in role_usage.items():
            ET.SubElement(role_elem, 'role', type=role, count=str(count))
        
        # Average elements per slide
        total_elements = sum(int(elem.text or 0) for elem in slides.findall('.//computed_features/element_count'))
        avg_elements = total_elements / slide_count if slide_count > 0 else 0
        ET.SubElement(stats, 'avg_elements_per_slide').text = f"{avg_elements:.1f}"
        
        # Color palette usage
        all_colors = []
        for color in training_data.findall('.//color[@hex]'):
            hex_val = color.get('hex')
            if hex_val and not hex_val.startswith('scheme:'):
                all_colors.append(hex_val)
        
        color_counts = defaultdict(int)
        for color in all_colors:
            color_counts[color] += 1
        
        top_colors = sorted(color_counts.items(), key=lambda x: x[1], reverse=True)[:10]
        colors_elem = ET.SubElement(stats, 'most_used_colors')
        total_color_count = len(all_colors)
        for color, count in top_colors:
            percentage = (count / total_color_count * 100) if total_color_count > 0 else 0
            ET.SubElement(colors_elem, 'color',
                         hex=color,
                         count=str(count),
                         percentage=f"{percentage:.1f}")
        
        return stats
    
    # Utility methods
    
    def get_relationship_target(self, rels_file, rel_id):
        """Get target file from relationship ID"""
        try:
            rels_xml = self.zip_file.read(rels_file)
            rels_root = ET.fromstring(rels_xml)
            
            for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                if rel.get('Id') == rel_id:
                    return rel.get('Target')
        except:
            pass
        return None
    
    def get_all_relationships(self, rels_file, rel_type):
        """Get all relationship targets of a specific type"""
        targets = []
        try:
            rels_xml = self.zip_file.read(rels_file)
            rels_root = ET.fromstring(rels_xml)
            
            for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                if rel_type in rel.get('Type', ''):
                    targets.append(rel.get('Target'))
        except:
            pass
        return targets
    
    def has_relationship_type(self, rels_file, rel_type):
        """Check if relationship type exists"""
        try:
            rels_xml = self.zip_file.read(rels_file)
            rels_root = ET.fromstring(rels_xml)
            
            for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                if rel_type in rel.get('Type', ''):
                    return True
        except:
            pass
        return False
    
    def get_text(self, root, xpath, default=''):
        """Get text from element via xpath"""
        elem = root.find(xpath, self.namespaces)
        return elem.text if elem is not None and elem.text else default
    
    def get_text_from_chart_element(self, elem):
        """Extract text from chart title/label elements"""
        if elem is None:
            return None
        
        # Try rich text
        rich = elem.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}t', self.namespaces)
        if rich is not None and rich.text:
            return rich.text
        
        # Try string reference
        str_ref = elem.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}strRef')
        if str_ref is not None:
            str_cache = str_ref.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}strCache')
            if str_cache is not None:
                pt = str_cache.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}pt')
                if pt is not None:
                    v = pt.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}v')
                    if v is not None and v.text:
                        return v.text
        
        return None
    
    def get_all_text_from_element(self, elem):
        """Get all text content from an element"""
        texts = []
        for t in elem.findall('.//a:t', self.namespaces):
            if t.text:
                texts.append(t.text)
        return ' '.join(texts)
    
    def compute_file_hash(self):
        """Compute SHA256 hash of file"""
        sha256 = hashlib.sha256()
        with open(self.pptx_path, 'rb') as f:
            for chunk in iter(lambda: f.read(4096), b''):
                sha256.update(chunk)
        return sha256.hexdigest()[:16]
    
    def compute_slide_hash(self, slide_xml):
        """Compute hash of slide content for deduplication"""
        sha256 = hashlib.sha256()
        sha256.update(slide_xml)
        return sha256.hexdigest()[:16]
    
    def generate_uuid(self, seed):
        """Generate deterministic UUID from seed"""
        import uuid
        return str(uuid.uuid5(uuid.NAMESPACE_DNS, seed))
    
    def calculate_aspect_ratio(self, width, height):
        """Calculate aspect ratio string"""
        from math import gcd
        divisor = gcd(width, height)
        w = width // divisor
        h = height // divisor
        return f"{w}:{h}"
    
    def hex_to_rgb(self, hex_color):
        """Convert hex color to RGB tuple"""
        if hex_color.startswith('#'):
            hex_color = hex_color[1:]
        if hex_color.startswith('scheme:') or hex_color.startswith('preset:'):
            return (128, 128, 128)  # Default gray for scheme colors
        try:
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            return (r, g, b)
        except:
            return (128, 128, 128)
    
    def rgb_to_lab(self, rgb):
        """Convert RGB to LAB color space for perceptual analysis"""
        # Normalize RGB values
        r, g, b = [x / 255.0 for x in rgb]
        
        # Convert to linear RGB
        def linearize(c):
            if c > 0.04045:
                return ((c + 0.055) / 1.055) ** 2.4
            else:
                return c / 12.92
        
        r = linearize(r)
        g = linearize(g)
        b = linearize(b)
        
        # Convert to XYZ
        x = r * 0.4124564 + g * 0.3575761 + b * 0.1804375
        y = r * 0.2126729 + g * 0.7151522 + b * 0.0721750
        z = r * 0.0193339 + g * 0.1191920 + b * 0.9503041
        
        # Convert to LAB
        def f(t):
            delta = 6/29
            if t > delta**3:
                return t**(1/3)
            else:
                return t/(3*delta**2) + 4/29
        
        # D65 white point
        xn, yn, zn = 0.95047, 1.00000, 1.08883
        
        fx = f(x/xn)
        fy = f(y/yn)
        fz = f(z/zn)
        
        l = 116 * fy - 16
        a = 500 * (fx - fy)
        b_lab = 200 * (fy - fz)
        
        return (l, a, b_lab)
    
    def close(self):
        """Close the zip file"""
        if hasattr(self, 'zip_file'):
            self.zip_file.close()


def main():
    """Main entry point for the extractor"""
    import sys
    import os
    
    if len(sys.argv) < 2:
        print("Usage: python slide_extractor.py <input.pptx> [output_folder]")
        print("\nExtracts comprehensive features from PowerPoint files for ML training")
        print("\nExamples:")
        print("  python slide_extractor/slide_extractor.py data/presentation.pptx")
        print("  python slide_extractor/slide_extractor.py data/presentation.pptx output/")
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    # Handle relative paths from project root
    if not os.path.isabs(input_file):
        # Get the project root (parent of slide_extractor folder)
        script_dir = Path(__file__).parent
        project_root = script_dir.parent
        input_file = project_root / input_file
    else:
        input_file = Path(input_file)
    
    if not input_file.exists():
        print(f"Error: File '{input_file}' not found")
        print(f"Looking in: {input_file.absolute()}")
        sys.exit(1)
    
    # Determine output location
    if len(sys.argv) > 2:
        output_location = Path(sys.argv[2])
    else:
        # Default: create 'output' folder in project root
        script_dir = Path(__file__).parent
        project_root = script_dir.parent
        output_location = project_root / 'output'
    
    # Create output directory if it doesn't exist
    output_location.mkdir(parents=True, exist_ok=True)
    
    # Generate output filename
    output_file = output_location / f"{input_file.stem}_training.xml"
    
    print(f"\n{'='*60}")
    print(f"PowerPoint Feature Extractor for ML Training")
    print(f"{'='*60}\n")
    
    try:
        extractor = PPTXFeatureExtractor(input_file, output_file)
        training_data = extractor.extract_all_features()
        extractor.close()
        
        print(f"\n{'='*60}")
        print(f"✓ Extraction complete!")
        print(f"  Output: {output_file}")
        print(f"  Format: Structured XML training data")
        print(f"{'='*60}\n")
        
        # Print summary
        slides = training_data.find('.//slides')
        if slides is not None:
            slide_count = len(slides.findall('.//slide'))
            print(f"Summary:")
            print(f"  • Total slides: {slide_count}")
            
            stats = training_data.find('.//global_statistics')
            if stats is not None:
                avg_elem = stats.find('.//avg_elements_per_slide')
                if avg_elem is not None:
                    print(f"  • Avg elements/slide: {avg_elem.text}")
                
                colors_elem = stats.find('.//most_used_colors')
                if colors_elem is not None:
                    color_count = len(colors_elem.findall('.//color'))
                    print(f"  • Unique colors: {color_count}")
        
        print(f"\nThe extracted data can now be used for:")
        print(f"  • Training generative models")
        print(f"  • Style transfer learning")
        print(f"  • Layout analysis")
        print(f"  • Design pattern recognition")
        
    except Exception as e:
        print(f"\n✗ Error during extraction: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()