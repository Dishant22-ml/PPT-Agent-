"""
Hybrid AI Slide Analyzer & Modifier - FULL FIXED VERSION
Combines mathematical analysis with LLM reasoning
Shows old and new values on modification
"""

import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import List, Dict, Optional, Tuple
import json
import re
from groq import Groq
import os


@dataclass
class SlideElement:
    """Represents a single element on a slide"""
    id: str
    type: str
    bounds: Dict[str, float]
    z_order: int
    text_content: str
    text_length: int
    word_count: int
    has_text: bool
    placeholder_type: Optional[str]
    # Mathematical analysis
    position_category: str
    horizontal_category: str
    size_category: str
    math_confidence: float
    # LLM analysis
    llm_category: Optional[str] = None
    llm_role: Optional[str] = None
    llm_confidence: Optional[float] = None
    llm_reasoning: Optional[str] = None
    # Final consensus
    final_category: Optional[str] = None
    final_confidence: Optional[float] = None


class HybridSlideAnalyzer:
    """Hybrid analyzer combining math + LLM intelligence"""
    
    def __init__(self, api_key: str = None):
        self.elements: List[SlideElement] = []
        self.slide_width = 9144000
        self.slide_height = 6858000
        self.client = Groq(api_key=api_key or os.getenv("GROQ_API_KEY"))
        self.model = "llama-3.3-70b-versatile"
    
    def analyze_xml(self, xml_path: str) -> Dict:
        """
        Complete hybrid analysis: Math + LLM
        Returns comprehensive understanding of slide
        """
        print("\n" + "="*80)
        print("🚀 HYBRID SLIDE ANALYSIS - Math + AI")
        print("="*80 + "\n")
        
        tree = ET.parse(xml_path)
        root = tree.getroot()
        
        # Extract slide dimensions
        slide = root.find('.//slide')
        if slide is not None:
            self.slide_width = int(slide.get('width', 9144000))
            self.slide_height = int(slide.get('height', 6858000))
        
        print("📐 PHASE 1: Mathematical Analysis")
        print("-" * 80)
        
        # Extract and categorize elements mathematically
        self.elements = []
        elements_node = root.find('.//elements')
        if elements_node is not None:
            for elem in elements_node.findall('.//element'):
                slide_elem = self._extract_element(elem)
                if slide_elem:
                    self.elements.append(slide_elem)
        
        # Also extract from shapes (which may contain text boxes)
        shapes_node = root.find('.//shapes')
        if shapes_node is not None:
            for shape in shapes_node.findall('.//shape'):
                slide_elem = self._extract_element(shape)
                if slide_elem:
                    self.elements.append(slide_elem)
        
        self._mathematical_categorization()
        print(f"✓ Analyzed {len(self.elements)} elements mathematically\n")
        
        print("🧠 PHASE 2: LLM Semantic Analysis")
        print("-" * 80)
        
        # Get LLM understanding
        self._llm_analysis()
        print("✓ LLM analysis complete\n")
        
        print("🔗 PHASE 3: Consensus Fusion")
        print("-" * 80)
        
        # Combine insights
        self._fuse_analyses()
        print("✓ Created unified understanding\n")
        
        # Build final analysis
        analysis = self._build_comprehensive_analysis()
        
        self._print_analysis(analysis)
        
        return analysis
    
    def _extract_element(self, elem: ET.Element) -> Optional[SlideElement]:
        """Extract element with all properties"""
        elem_id = elem.get('id')
        
        # Skip elements without ID
        if not elem_id:
            return None
            
        elem_type = elem.get('type', 'unknown')
        z_order = int(elem.get('z_order', 0))
        
        geom = elem.find('.//geometry')
        if geom is None:
            return None
        
        x = float(geom.findtext('.//x', '0'))
        y = float(geom.findtext('.//y', '0'))
        width = float(geom.findtext('.//width', '0'))
        height = float(geom.findtext('.//height', '0'))
        
        text_content = self._extract_all_text(elem)
        
        placeholder = elem.find('.//placeholder')
        placeholder_type = placeholder.get('type') if placeholder is not None else None
        
        return SlideElement(
            id=elem_id,
            type=elem_type,
            bounds={'x': x, 'y': y, 'width': width, 'height': height},
            z_order=z_order,
            text_content=text_content,
            text_length=len(text_content),
            word_count=len(text_content.split()),
            has_text=bool(text_content.strip()),
            placeholder_type=placeholder_type,
            position_category='unknown',
            horizontal_category='unknown',
            size_category='unknown',
            math_confidence=0.0
        )
    
    def _extract_all_text(self, elem: ET.Element) -> str:
        """Extract text from all XML structures - including nested"""
        texts = []
        
        # text_body structure
        text_body = elem.find('.//text_body')
        if text_body is not None:
            for para in text_body.findall('.//paragraph'):
                for text_elem in para.findall('.//text'):
                    if text_elem.text:
                        texts.append(text_elem.text.strip())
        
        # text_run structure
        for text_run in elem.findall('.//text_run'):
            text_elem = text_run.find('.//text')
            if text_elem is not None and text_elem.text:
                texts.append(text_elem.text.strip())
        
        # Direct text elements
        for text_elem in elem.findall('.//text'):
            if text_elem.text and text_elem.text.strip():
                texts.append(text_elem.text.strip())
        
        return ' '.join(texts)
    
    def _mathematical_categorization(self):
        """Mathematical/geometric categorization"""
        for elem in self.elements:
            elem.position_category = self._get_position_category(elem.bounds['y'], elem.bounds['height'])
            elem.horizontal_category = self._get_horizontal_category(elem.bounds['x'], elem.bounds['width'])
            elem.size_category = self._get_size_category(elem.bounds['width'] * elem.bounds['height'])
            elem.math_confidence = self._calculate_math_confidence(elem)
    
    def _get_position_category(self, y: float, height: float) -> str:
        center_y = y + height / 2
        normalized = center_y / self.slide_height if self.slide_height > 0 else 0.5
        if normalized < 0.2:
            return 'top'
        elif normalized < 0.4:
            return 'upper-mid'
        elif normalized < 0.6:
            return 'mid'
        elif normalized < 0.8:
            return 'lower-mid'
        else:
            return 'bottom'
    
    def _get_horizontal_category(self, x: float, width: float) -> str:
        center_x = x + width / 2
        normalized = center_x / self.slide_width if self.slide_width > 0 else 0.5
        if normalized < 0.33:
            return 'L'
        elif normalized < 0.67:
            return 'C'
        else:
            return 'R'
    
    def _get_size_category(self, area: float) -> str:
        total_area = self.slide_width * self.slide_height
        normalized_area = area / total_area if total_area > 0 else 0
        if normalized_area < 0.01:
            return 'XS'
        elif normalized_area < 0.05:
            return 'S'
        elif normalized_area < 0.15:
            return 'M'
        elif normalized_area < 0.4:
            return 'L'
        else:
            return 'XL'
    
    def _calculate_math_confidence(self, elem: SlideElement) -> float:
        """Calculate confidence based on mathematical features"""
        confidence = 0.5
        
        if elem.placeholder_type in ['title', 'ctrTitle']:
            confidence = 0.95
        elif elem.position_category == 'top' and elem.has_text and elem.size_category in ['M', 'L']:
            if elem.word_count < 15:
                confidence = 0.75
        elif elem.position_category in ['mid', 'lower-mid'] and elem.has_text:
            if elem.word_count > 10:
                confidence = 0.70
        
        return confidence
    
    def _sanitize_json_string(self, s: str) -> str:
        """Sanitize and repair malformed JSON"""
        # Remove control characters
        s = ''.join(ch for ch in s if ord(ch) >= 32 or ch in '\n\r\t')
        # Escape unescaped quotes within strings (basic attempt)
        s = s.replace('\\', '\\\\')
        s = s.replace('\\"', '"')
        return s
    
    def _parse_json_safely(self, response_text: str) -> Optional[Dict]:
        """Parse JSON with multiple fallback strategies"""
        
        # Strategy 1: Extract JSON from code blocks
        if "```json" in response_text:
            try:
                json_str = response_text.split("```json")[1].split("```")[0].strip()
                return json.loads(json_str)
            except json.JSONDecodeError:
                pass
        
        if "```" in response_text:
            try:
                json_str = response_text.split("```")[1].split("```")[0].strip()
                return json.loads(json_str)
            except json.JSONDecodeError:
                pass
        
        # Strategy 2: Find JSON object in response
        match = re.search(r'\{.*\}', response_text, re.DOTALL)
        if match:
            try:
                return json.loads(match.group())
            except json.JSONDecodeError:
                pass
        
        # Strategy 3: Sanitize and retry
        try:
            sanitized = self._sanitize_json_string(response_text)
            return json.loads(sanitized)
        except json.JSONDecodeError:
            pass
        
        # Strategy 4: Extract just the elements we need
        try:
            elements_match = re.search(r'"elements"\s*:\s*\{(.*?)\}(?:\s*[,}])', response_text, re.DOTALL)
            if elements_match:
                return {"elements": {}, "overall": "Partial parse"}
        except:
            pass
        
        return None
    
    def _llm_analysis(self):
        """Use LLM to understand slide semantically - OPTIMIZED FOR TOKENS"""
        
        context = self._build_compact_context()
        
        system_prompt = """Analyze slide elements. For each, determine:
- category: title|subtitle|body|image|chart|decoration
- role: brief purpose (max 20 chars)
- confidence: 0-1

Rules:
- TITLE: top, short (<15w), prominent
- SUBTITLE: below title  
- BODY: middle, longer text
- Use text content

STRICT JSON FORMAT:
{"overall":"brief analysis","elements":{"id":{"category":"title","role":"main","confidence":0.95,"reasoning":"why"}}}"""

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": context}
                ],
                temperature=0.3,
                max_tokens=1500
            )
            
            response_text = response.choices[0].message.content
            print(f"📨 Raw response length: {len(response_text)} chars")
            
            # Use robust JSON parsing
            llm_result = self._parse_json_safely(response_text)
            
            if llm_result is None:
                print("⚠️  Could not parse JSON response")
                print("📐 Using mathematical analysis only\n")
                return None
            
            overall = llm_result.get('overall', 'Analyzed')
            print(f"📊 LLM: {overall[:60]}{'...' if len(overall) > 60 else ''}\n")
            
            # Apply LLM insights to elements
            element_analyses = llm_result.get('elements', {})
            for elem in self.elements:
                if elem.id in element_analyses:
                    analysis = element_analyses[elem.id]
                    elem.llm_category = analysis.get('category')
                    elem.llm_role = analysis.get('role', '')[:50]
                    elem.llm_confidence = analysis.get('confidence', 0.5)
                    elem.llm_reasoning = analysis.get('reasoning', '')[:80]
            
            return llm_result
            
        except Exception as e:
            print(f"⚠️  LLM analysis failed: {str(e)[:100]}")
            print("📐 Using mathematical analysis only\n")
            return None
    
    def _build_compact_context(self) -> str:
        """Build ULTRA-COMPACT context to save tokens"""
        sorted_elems = sorted(self.elements, key=lambda e: e.bounds['y'])
        
        important_elems = [e for e in sorted_elems if e.has_text or e.bounds['width'] * e.bounds['height'] > 0.05][:20]
        
        context = "ELEMENTS:\n"
        
        for i, elem in enumerate(important_elems, 1):
            area = elem.bounds['width'] * elem.bounds['height']
            elem_id_display = elem.id[:15] if elem.id else f"elem_{i}"
            context += f"{i}.{elem_id_display} "
            context += f"T:{elem.type} "
            context += f"P:{elem.position_category}-{elem.horizontal_category} "
            context += f"S:{elem.size_category} "
            
            if elem.placeholder_type:
                context += f"PH:{elem.placeholder_type} "
            
            if elem.has_text:
                preview = elem.text_content[:60].replace('\n', ' ').replace('"', "'")
                context += f'Txt({elem.word_count}w):"{preview}"'
            else:
                context += "(visual)"
            
            context += "\n"
        
        return context
    
    def _fuse_analyses(self):
        """Combine mathematical and LLM analyses intelligently"""
        for elem in self.elements:
            if elem.llm_category and elem.llm_confidence:
                combined_confidence = (0.6 * elem.llm_confidence) + (0.4 * elem.math_confidence)
                
                math_category = self._infer_math_category(elem)
                if math_category == elem.llm_category:
                    combined_confidence = min(0.98, combined_confidence * 1.2)
                
                elem.final_category = elem.llm_category
                elem.final_confidence = combined_confidence
            else:
                elem.final_category = self._infer_math_category(elem)
                elem.final_confidence = elem.math_confidence
    
    def _infer_math_category(self, elem: SlideElement) -> str:
        """Infer category from mathematical features"""
        if elem.placeholder_type in ['title', 'ctrTitle']:
            return 'title'
        elif elem.placeholder_type == 'subTitle':
            return 'subtitle'
        elif elem.position_category == 'top' and elem.has_text and elem.size_category in ['M', 'L']:
            if elem.word_count < 15:
                return 'title'
        elif elem.position_category in ['mid', 'lower-mid'] and elem.has_text:
            return 'body'
        elif not elem.has_text and elem.bounds['width'] * elem.bounds['height'] > 0.1:
            return 'image'
        return 'content'
    
    def _build_comprehensive_analysis(self) -> Dict:
        """Build final comprehensive analysis"""
        sorted_elements = sorted(self.elements, key=lambda e: e.bounds['y'])
        
        title_elem = next((e for e in sorted_elements if e.final_category == 'title'), None)
        subtitle_elem = next((e for e in sorted_elements if e.final_category == 'subtitle'), None)
        body_elems = [e for e in sorted_elements if e.final_category == 'body']
        image_elems = [e for e in sorted_elements if e.final_category in ['image', 'chart']]
        
        return {
            'slide_dimensions': {
                'width': self.slide_width,
                'height': self.slide_height
            },
            'elements': [self._serialize_element(e) for e in sorted_elements],
            'categorized': {
                'title': self._serialize_element(title_elem) if title_elem else None,
                'subtitle': self._serialize_element(subtitle_elem) if subtitle_elem else None,
                'body': [self._serialize_element(e) for e in body_elems],
                'images': [self._serialize_element(e) for e in image_elems]
            },
            'statistics': {
                'total_elements': len(self.elements),
                'text_elements': len([e for e in self.elements if e.has_text]),
                'avg_confidence': sum(e.final_confidence or 0 for e in self.elements) / len(self.elements) if self.elements else 0,
                'llm_analyzed': sum(1 for e in self.elements if e.llm_category is not None)
            }
        }
    
    def _serialize_element(self, elem: SlideElement) -> Optional[Dict]:
        """Convert element to dictionary"""
        if elem is None:
            return None
        
        return {
            'id': elem.id,
            'type': elem.type,
            'bounds': elem.bounds,
            'position': f"{elem.position_category} {elem.horizontal_category}",
            'size': elem.size_category,
            'category': elem.final_category,
            'confidence': round(elem.final_confidence or 0, 2),
            'text': elem.text_content[:200] if elem.text_content else '',
            'text_length': elem.text_length,
            'word_count': elem.word_count,
            'z_order': elem.z_order,
            'llm_role': elem.llm_role,
            'llm_reasoning': elem.llm_reasoning,
            'math_confidence': round(elem.math_confidence, 2),
            'llm_confidence': round(elem.llm_confidence, 2) if elem.llm_confidence else None
        }
    
    def _print_analysis(self, analysis: Dict):
        """Print beautiful analysis summary"""
        print("\n" + "="*80)
        print("🎯 FINAL HYBRID ANALYSIS RESULTS")
        print("="*80 + "\n")
        
        stats = analysis['statistics']
        print(f"📊 Statistics:")
        print(f"   Total Elements: {stats['total_elements']}")
        print(f"   Text Elements: {stats['text_elements']}")
        print(f"   LLM Analyzed: {stats['llm_analyzed']}/{stats['total_elements']}")
        print(f"   Average Confidence: {stats['avg_confidence']:.1%}\n")
        
        cat = analysis['categorized']
        
        if cat['title']:
            t = cat['title']
            print(f"📌 TITLE (Confidence: {t['confidence']:.0%})")
            print(f"   ID: {t['id']}")
            print(f"   Position: {t['position']}")
            print(f"   Text: \"{t['text']}\"")
            if t.get('llm_role'):
                print(f"   LLM Role: {t['llm_role']}")
            if t.get('llm_reasoning'):
                print(f"   LLM Reasoning: {t['llm_reasoning']}")
            print()
        
        if cat['subtitle']:
            s = cat['subtitle']
            print(f"📌 SUBTITLE (Confidence: {s['confidence']:.0%})")
            print(f"   ID: {s['id']}")
            print(f"   Text: \"{s['text']}\"")
            if s.get('llm_role'):
                print(f"   LLM Role: {s['llm_role']}")
            print()
        
        if cat['body']:
            print(f"📝 BODY ELEMENTS ({len(cat['body'])})")
            for i, body in enumerate(cat['body'][:3], 1):
                print(f"   {i}. {body['id']} (Confidence: {body['confidence']:.0%})")
                print(f"      Text: \"{body['text'][:100]}...\"")
                if body.get('llm_role'):
                    print(f"      Role: {body['llm_role']}")
                print()
        
        if cat['images']:
            print(f"🖼️  IMAGES/CHARTS ({len(cat['images'])})")
            for i, img in enumerate(cat['images'][:3], 1):
                print(f"   {i}. {img['id']} - {img['type']} ({img['size']})")
                if img.get('llm_role'):
                    print(f"      Role: {img['llm_role']}")
                print()
        
        print("="*80 + "\n")


class IntelligentSlideModifier:
    """Intelligent modifier using hybrid analysis"""
    
    def __init__(self, api_key: str = None):
        self.client = Groq(api_key=api_key or os.getenv("GROQ_API_KEY"))
        self.model = "llama-3.3-70b-versatile"
    
    def modify_slide(self, xml_path: str, analysis: Dict, prompt: str) -> str:
        """Modify slide intelligently"""
        print("\n" + "="*80)
        print("🎨 INTELLIGENT MODIFICATION")
        print("="*80)
        print(f"Request: {prompt}\n")
        
        modifications = self._get_modifications(analysis, prompt)
        
        if not modifications:
            print("❌ No modifications determined\n")
            return xml_path
        
        output_path = self._apply_modifications(xml_path, modifications)
        
        print(f"\n✅ Modified slide: {output_path}")
        print("="*80 + "\n")
        
        return output_path
    
    def _get_modifications(self, analysis: Dict, prompt: str) -> List[Dict]:
        """Determine modifications using analysis - COMPREHENSIVE"""
        
        context = self._build_compact_modification_context(analysis)
        
        system_prompt = """You are a comprehensive slide modification assistant. Your task is to identify ALL text content that needs modification, including bullet points, nested text, and multi-part segments.

CRITICAL RULES:
1. Use EXACT element IDs from [id=...] ONLY
2. Find and modify ALL content about the old topic: main text, bullets, sub-points, details
3. Each element may contain multiple sentences/segments - replace ENTIRE content, not just keywords
4. Include elements with keywords: Graphite, Energy Transition, mineral, India reserves, India players, etc.
5. For each element, provide a complete replacement that maintains structure but changes topic

JSON FORMAT (STRICT):
{"analysis":"summary of changes","modifications":[{"element_id":"ID","action":"replace_text","old_value":"full original text","new_value":"complete new text","confidence":0.9,"reasoning":"why"}]}

KEY INSTRUCTIONS:
- If element contains "Graphite" + "mining", replace with "AI in mining"
- If element contains "India reserves", replace with "AI-powered resource management" 
- If element contains bullet points like "Point 1. Text. Point 2. Text. Point 3. Text.", replace all with new content maintaining structure
- Look at word count in "Segments" field - high count means multi-part content that needs full replacement
- Be exhaustive - find 15-30+ modifications if content is multi-sectioned

EXAMPLE:
Old: "The steel industry uses graphite. Graphite is essential for batteries. EVs need graphite."
New: "AI is transforming resource management. AI enables better forecasting. Smart systems optimize allocation."
"""

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"{context}\n\nREQUEST: {prompt}"}
                ],
                temperature=0.35,
                max_tokens=4000
            )
            
            response_text = response.choices[0].message.content
            
            # Use same robust parsing
            result = self._parse_json_safely(response_text)
            
            if result is None:
                print("❌ Could not parse LLM response")
                return []
            
            print(f"🎯 {result.get('analysis', 'N/A')}\n")
            print("📝 Modifications:")
            for i, mod in enumerate(result.get('modifications', []), 1):
                elem_id = mod['element_id']
                old_val = mod.get('old_value', 'N/A')[:60]
                new_val = mod.get('new_value', 'N/A')[:60]
                print(f"   {i}. Element ID: {elem_id}")
                print(f"      Old: {old_val}")
                print(f"      New: {new_val}")
                print(f"      Confidence: {mod.get('confidence', 0):.0%}")
                print(f"      Reason: {mod.get('reasoning', '')[:60]}\n")
            
            return result.get('modifications', [])
            
        except Exception as e:
            print(f"❌ Error: {str(e)[:100]}")
            return []
    
    def _parse_json_safely(self, response_text: str) -> Optional[Dict]:
        """Parse JSON with multiple fallback strategies"""
        
        # Strategy 1: Extract JSON from code blocks
        if "```json" in response_text:
            try:
                json_str = response_text.split("```json")[1].split("```")[0].strip()
                return json.loads(json_str)
            except json.JSONDecodeError:
                pass
        
        if "```" in response_text:
            try:
                json_str = response_text.split("```")[1].split("```")[0].strip()
                return json.loads(json_str)
            except json.JSONDecodeError:
                pass
        
        # Strategy 2: Find JSON object in response
        match = re.search(r'\{.*\}', response_text, re.DOTALL)
        if match:
            try:
                return json.loads(match.group())
            except json.JSONDecodeError:
                pass
        
        return None
    
    def _build_compact_modification_context(self, analysis: Dict) -> str:
        """Build comprehensive context with ALL elements including shapes with text"""
        context = "ALL SLIDE ELEMENTS WITH TEXT:\n"
        context += "(Including text boxes, shapes, paragraphs, bullets, and nested content)\n\n"
        
        # Add statistics
        context += f"Total Elements: {analysis['statistics']['total_elements']}\n"
        context += f"Text Elements: {analysis['statistics']['text_elements']}\n\n"
        
        # List ALL elements that have text content
        context += "ELEMENTS WITH TEXT CONTENT (elements + shapes):\n"
        text_elements = [e for e in analysis['elements'] if e['text'] and e['text'].strip()]
        
        for i, elem in enumerate(text_elements, 1):
            elem_id = elem['id']
            elem_type = elem['type']
            category = elem['category']
            pos = elem['position']
            size = elem['size']
            text = elem['text'][:100] if elem['text'] else "(no text)"
            
            # Count sentences/segments in text
            segments = [s.strip() for s in text.split('.') if s.strip()]
            segment_count = len(segments)
            
            context += f"{i}. [id={elem_id}] type={elem_type} category={category}\n"
            context += f"   Position: {pos} | Size: {size} | Segments: {segment_count}\n"
            context += f"   Text: \"{text}\"\n\n"
        
        context += f"Total text-bearing elements: {len(text_elements)}\n\n"
        context += "IMPORTANT NOTES:\n"
        context += "- Elements include text boxes, shapes with text, paragraphs, bullets, and all nested content\n"
        context += "- Each element may contain MULTIPLE sentences, bullet points, or segments\n"
        context += "- Modify ALL text within an element, not just the main sentence\n"
        context += "- Look for 'Graphite', 'Energy Transition', 'India', 'mineral' keywords and replace with AI-related content\n"
        context += "- Include elements with partial matches of old keywords\n\n"
        
        cat = analysis['categorized']
        
        context += "KEY IDENTIFIED ELEMENTS:\n"
        if cat['title']:
            t = cat['title']
            context += f"- TITLE [id={t['id']}]: \"{t['text'][:80]}\"\n"
        
        if cat['subtitle']:
            s = cat['subtitle']
            context += f"- SUBTITLE [id={s['id']}]: \"{s['text'][:80]}\"\n"
        
        if cat['body']:
            context += f"- BODY ELEMENTS ({len(cat['body'])} identified):\n"
            for b in cat['body']:
                context += f"  [id={b['id']}] \"{b['text'][:60]}\"\n"
        
        if cat['images']:
            context += f"- VISUAL ELEMENTS ({len(cat['images'])} - charts, images, etc.):\n"
            for img in cat['images'][:5]:
                context += f"  [id={img['id']}] type={img['type']} ({img['size']})\n"
        
        context += "\n⚠️  IMPORTANT: Use EXACT id values from [id=...] in your response.\n"
        context += "Consider ALL relevant elements (text boxes, labels, titles, etc.) for modification.\n"
        
        return context
    
    def _extract_element_text(self, elem: ET.Element) -> str:
        """Extract current text from element"""
        texts = []
        
        text_body = elem.find('.//text_body')
        if text_body is not None:
            for para in text_body.findall('.//paragraph'):
                for text_elem in para.findall('.//text'):
                    if text_elem.text:
                        texts.append(text_elem.text.strip())
        
        for text_run in elem.findall('.//text_run'):
            text_elem = text_run.find('.//text')
            if text_elem is not None and text_elem.text:
                texts.append(text_elem.text.strip())
        
        for text_elem in elem.findall('.//text'):
            if text_elem.text and text_elem.text.strip():
                texts.append(text_elem.text.strip())
        
        current = ' '.join(texts)
        return current[:100] if current else "(empty)"
    
    def _apply_modifications(self, xml_path: str, modifications: List[Dict]) -> str:
        """Apply modifications to XML with smart fallbacks"""
        tree = ET.parse(xml_path)
        root = tree.getroot()
        
        successful = 0
        failed = 0
        
        # Debug: List all available element IDs
        all_elements = root.findall('.//element')
        all_ids = [e.get('id') for e in all_elements]
        print(f"   Available element IDs: {all_ids}\n")
        
        for mod in modifications:
            elem_id = mod['element_id']
            
            # Strip "id=" prefix if present (LLM sometimes includes it)
            if isinstance(elem_id, str) and elem_id.startswith('id='):
                elem_id = elem_id.replace('id=', '')
            
            action = mod['action']
            new_value = mod['new_value']
            
            # Strategy 1: Direct attribute match
            elem = root.find(f".//element[@id='{elem_id}']")
            
            # Strategy 2: Try string conversion if numeric
            if elem is None and str(elem_id).isdigit():
                elem = root.find(f".//element[@id='{str(elem_id)}']")
            
            # Strategy 3: Linear search as fallback
            if elem is None:
                for e in all_elements:
                    if e.get('id') == str(elem_id) or e.get('id') == elem_id:
                        elem = e
                        break
            
            if elem is None:
                print(f"⚠️  Element ID '{elem_id}' not found in {all_ids}")
                failed += 1
                continue
            
            if action == "replace_text":
                print(f"      🔍 Attempting to find element '{elem_id}'...")
                old_text = self._extract_element_text(elem)
                success = self._replace_text_smart(elem, new_value)
                if success:
                    print(f"✅ Updated {elem_id}")
                    print(f"      Old: '{old_text}'")
                    print(f"      New: '{new_value}'")
                    successful += 1
                else:
                    print(f"❌ Failed to modify {elem_id}")
                    print(f"      Old: '{old_text}'")
                    failed += 1
        
        print(f"\n📊 {successful} successful, {failed} failed")
        
        output_path = xml_path.replace('.xml', '_modified.xml')
        tree.write(output_path, encoding='utf-8', xml_declaration=True)
        
        return output_path
    
    def _replace_text_smart(self, elem: ET.Element, new_text: str) -> bool:
        """Smart text replacement with fallbacks"""
        
        # Strategy 1: text_body > paragraph > text (most common in PPTX)
        text_body = elem.find('.//text_body')
        if text_body is not None:
            paragraphs = text_body.findall('.//paragraph')
            if paragraphs:
                # Replace text in first paragraph
                for paragraph in paragraphs:
                    text_elems = paragraph.findall('.//text')
                    if text_elems:
                        text_elems[0].text = new_text
                        # Clear other text elements in this paragraph
                        for te in text_elems[1:]:
                            te.text = ""
                        return True
            
            # Create new paragraph if needed
            paragraph = ET.SubElement(text_body, 'paragraph')
            text_run = ET.SubElement(paragraph, 'text_run')
            text_elem = ET.SubElement(text_run, 'text')
            text_elem.text = new_text
            return True
        
        # Strategy 2: Direct text elements
        text_elems = elem.findall('.//text')
        if text_elems:
            text_elems[0].text = new_text
            # Clear other text elements
            for te in text_elems[1:]:
                te.text = ""
            return True
        
        # Strategy 3: text_run structure
        text_run = elem.find('.//text_run')
        if text_run is not None:
            text_elem = text_run.find('text')
            if text_elem is None:
                text_elem = ET.SubElement(text_run, 'text')
            text_elem.text = new_text
            return True
        
        # Strategy 4: Create full structure
        try:
            text_body = ET.SubElement(elem, 'text_body')
            paragraph = ET.SubElement(text_body, 'paragraph')
            text_run = ET.SubElement(paragraph, 'text_run')
            text_elem = ET.SubElement(text_run, 'text')
            text_elem.text = new_text
            print(f"      ℹ️  Created new text structure")
            return True
        except Exception as e:
            print(f"      ❌ Failed to create structure: {e}")
            return False


def main():
    """Main execution"""
    import sys
    
    if len(sys.argv) < 3:
        print("\nUsage: python hybrid_analyzer.py <xml_file> <groq_api_key> [prompt]")
        print("\nExample:")
        print('  python hybrid_analyzer.py slide.xml gsk_xxx')
        print('  python hybrid_analyzer.py slide.xml gsk_xxx "Change title to AI Revolution 2024"')
        return
    
    xml_file = sys.argv[1]
    api_key = sys.argv[2]
    prompt = sys.argv[3] if len(sys.argv) > 3 else None
    
    if not os.path.exists(xml_file):
        print(f"❌ File not found: {xml_file}")
        return
    
    # Analyze
    analyzer = HybridSlideAnalyzer(api_key=api_key)
    analysis = analyzer.analyze_xml(xml_file)
    
    # Save analysis
    with open('hybrid_analysis.json', 'w') as f:
        json.dump(analysis, f, indent=2)
    print("💾 Analysis saved to: hybrid_analysis.json\n")
    
    # Modify if prompt provided
    if prompt:
        modifier = IntelligentSlideModifier(api_key=api_key)
        output = modifier.modify_slide(xml_file, analysis, prompt)
        print(f"✅ Complete! Modified file: {output}")


if __name__ == "__main__":
    main()