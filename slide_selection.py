"""
Organization-Narrative-Driven Slide Template Selector (with Persistence)

Features:
- Organizational Storytelling Profile Creation
- Narrative Pattern Analysis using LLM
- Persistent Storage (no retraining needed)
- Style-Based Template Selection
"""

import xml.etree.ElementTree as ET
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Set
import json
from groq import Groq
import os
import re
import pickle
from datetime import datetime
import numpy as np


@dataclass
class NarrativePattern:
    """Represents a storytelling pattern"""
    pattern_name: str
    description: str
    opening_style: str
    flow_structure: str
    conclusion_style: str
    typical_layout: str
    visual_approach: str
    tone: str
    keywords: Set[str]
    frequency_in_org: float
    effectiveness_score: float


@dataclass
class SlideProfile:
    """Complete profile of a slide"""
    slide_id: str
    slide_index: int
    raw_text: str
    narrative_analysis: Dict
    keywords: Set[str]
    detected_patterns: List[Tuple[str, float]]
    semantic_role: str


class OrgNarrativeProfiler:
    """Build and persist organizational narrative DNA"""
    
    def __init__(self, groq_client, profile_dir: Path = Path("org_profiles")):
        self.groq_client = groq_client
        self.profile_dir = profile_dir
        self.profile_dir.mkdir(exist_ok=True)
    
    def profile_exists(self, org_name: str) -> bool:
        """Check if org profile already exists"""
        profile_file = self.profile_dir / f"{org_name}_profile.pkl"
        return profile_file.exists()
    
    def load_profile(self, org_name: str) -> Dict:
        """Load existing organizational profile"""
        profile_file = self.profile_dir / f"{org_name}_profile.pkl"
        
        if profile_file.exists():
            with open(profile_file, 'rb') as f:
                profile = pickle.load(f)
            print(f"✓ Loaded existing profile for {org_name}")
            print(f"  Created: {profile.get('created_at')}")
            print(f"  Slides analyzed: {profile.get('total_slides')}\n")
            return profile
        
        return None
    
    def save_profile(self, org_name: str, profile_data: Dict):
        """Save organizational profile to disk"""
        profile_file = self.profile_dir / f"{org_name}_profile.pkl"
        
        with open(profile_file, 'wb') as f:
            pickle.dump(profile_data, f)
        
        print(f"✓ Profile saved: {profile_file}\n")
    
    def build_profile(self, xml_path: Path, org_name: str) -> Dict:
        """Analyze slides and build organization profile"""
        
        print(f"\n{'='*80}")
        print(f"TRAINING ORGANIZATIONAL NARRATIVE MODEL")
        print(f"{'='*80}\n")
        
        # Load slides
        slides = self._load_slides(xml_path)
        print(f"✓ Loaded {len(slides)} slides\n")
        
        # Analyze each slide with LLM
        print("Analyzing slides with LLM (this may take a minute)...\n")
        slide_profiles = []
        
        for idx, slide in enumerate(slides, 1):
            print(f"  [{idx}/{len(slides)}] Analyzing slide {slide.slide_index}...", end="", flush=True)
            profile = self._analyze_slide_with_llm(slide)
            slide_profiles.append(profile)
            print(" ✓")
        
        print()
        
        # Extract organizational patterns
        print("Extracting organizational narrative patterns...")
        patterns = self._extract_patterns(slide_profiles)
        
        print(f"✓ Detected {len(patterns)} unique patterns\n")
        
        # Build profile
        profile = {
            'org_name': org_name,
            'created_at': datetime.now().isoformat(),
            'total_slides': len(slides),
            'slide_profiles': slide_profiles,
            'patterns': patterns,
            'storytelling_values': self._calculate_storytelling_values(slide_profiles),
            'keywords_frequency': self._calculate_keyword_frequency(slide_profiles)
        }
        
        return profile
    
    def _load_slides(self, xml_path: Path) -> List[SlideProfile]:
        """Load slides from XML"""
        tree = ET.parse(xml_path)
        root = tree.getroot()
        
        slides = []
        for slide_elem in root.findall('.//slide'):
            slide_id = slide_elem.get('id', 'unknown')
            slide_index = int(slide_elem.get('index', 0))
            
            # Extract text
            text_content = []
            for text_elem in slide_elem.findall('.//text'):
                if text_elem.text:
                    text_content.append(text_elem.text)
            
            raw_text = ' '.join(text_content)
            
            semantic_role = 'content'
            sem_elem = slide_elem.find('.//semantic_role')
            if sem_elem is not None and sem_elem.text:
                semantic_role = sem_elem.text
            
            slide = SlideProfile(
                slide_id=slide_id,
                slide_index=slide_index,
                raw_text=raw_text[:1000],  # Limit to 1000 chars
                narrative_analysis={},
                keywords=set(),
                detected_patterns=[],
                semantic_role=semantic_role
            )
            slides.append(slide)
        
        return slides
    
    def _analyze_slide_with_llm(self, slide: SlideProfile) -> SlideProfile:
        """Use LLM to analyze slide narrative"""
        
        if not slide.raw_text or len(slide.raw_text.strip()) < 20:
            slide.narrative_analysis = self._default_analysis()
            return slide
        
        prompt = f"""Analyze this slide's narrative structure and storytelling approach.

SLIDE CONTENT:
{slide.raw_text[:500]}

Provide analysis in JSON format:
{{
  "story_type": "problem-solution|comparison|journey|data-driven|narrative|other",
  "tone": "professional|conversational|technical|inspirational|critical|neutral",
  "opening": "hook|context|question|problem|statement",
  "flow": "linear|escalation|circular|contrast|buildup",
  "conclusion": "summary|call-to-action|insight|question|transition",
  "layout_style": "title-content|split|full-bleed|minimal|centered",
  "visual_approach": "data-heavy|narrative|minimalist|bold|balanced",
  "keywords": ["keyword1", "keyword2", "keyword3"],
  "narrative_strength": 0.0-1.0,
  "summary": "one sentence"
}}"""
        
        try:
            response = self.groq_client.chat.completions.create(
                model="mixtral-8x7b-32768",
                messages=[
                    {"role": "system", "content": "Analyze slide narrative. Return ONLY valid JSON."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,
                max_tokens=400
            )
            
            content = response.choices[0].message.content.strip()
            
            # Extract JSON
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            if json_match:
                analysis = json.loads(json_match.group())
                slide.narrative_analysis = analysis
                slide.keywords = set(analysis.get('keywords', []))
                return slide
        except Exception as e:
            pass
        
        slide.narrative_analysis = self._default_analysis()
        return slide
    
    def _default_analysis(self) -> Dict:
        """Fallback analysis"""
        return {
            "story_type": "neutral",
            "tone": "professional",
            "opening": "context",
            "flow": "linear",
            "conclusion": "summary",
            "layout_style": "title-content",
            "visual_approach": "balanced",
            "keywords": [],
            "narrative_strength": 0.5,
            "summary": "Standard slide"
        }
    
    def _extract_patterns(self, slide_profiles: List[SlideProfile]) -> List[NarrativePattern]:
        """Extract unique narrative patterns from slides"""
        
        patterns_dict = {}
        
        for slide in slide_profiles:
            analysis = slide.narrative_analysis
            key = f"{analysis.get('tone', 'neutral')}_{analysis.get('flow', 'linear')}"
            
            if key not in patterns_dict:
                patterns_dict[key] = {
                    'count': 0,
                    'analysis': analysis,
                    'keywords': set()
                }
            
            patterns_dict[key]['count'] += 1
            patterns_dict[key]['keywords'].update(analysis.get('keywords', []))
        
        patterns = []
        total = len(slide_profiles)
        
        for key, data in patterns_dict.items():
            analysis = data['analysis']
            freq = data['count'] / total
            
            pattern = NarrativePattern(
                pattern_name=f"{analysis.get('tone', 'neutral').title()} - {analysis.get('flow', 'linear').title()}",
                description=analysis.get('summary', ''),
                opening_style=analysis.get('opening', 'context'),
                flow_structure=analysis.get('flow', 'linear'),
                conclusion_style=analysis.get('conclusion', 'summary'),
                typical_layout=analysis.get('layout_style', 'title-content'),
                visual_approach=analysis.get('visual_approach', 'balanced'),
                tone=analysis.get('tone', 'professional'),
                keywords=data['keywords'],
                frequency_in_org=freq,
                effectiveness_score=0.7
            )
            patterns.append(pattern)
        
        return sorted(patterns, key=lambda x: x.frequency_in_org, reverse=True)
    
    def _calculate_storytelling_values(self, slide_profiles: List[SlideProfile]) -> Dict[str, float]:
        """Calculate org storytelling characteristics"""
        
        values = {
            "data_driven": 0.0,
            "narrative_heavy": 0.0,
            "technical": 0.0,
            "conversational": 0.0,
            "minimalist": 0.0
        }
        
        for slide in slide_profiles:
            analysis = slide.narrative_analysis
            visual = analysis.get('visual_approach', 'balanced').lower()
            tone = analysis.get('tone', 'professional').lower()
            
            if 'data' in visual or 'chart' in slide.raw_text.lower():
                values['data_driven'] += 0.1
            if 'narrative' in visual:
                values['narrative_heavy'] += 0.1
            if tone == 'technical':
                values['technical'] += 0.1
            if tone == 'conversational':
                values['conversational'] += 0.1
            if 'minimal' in visual:
                values['minimalist'] += 0.1
        
        # Normalize
        total = sum(values.values())
        if total > 0:
            values = {k: min(v / total, 1.0) for k, v in values.items()}
        
        return values
    
    def _calculate_keyword_frequency(self, slide_profiles: List[SlideProfile]) -> Dict[str, int]:
        """Extract most common keywords across org"""
        
        keyword_freq = {}
        for slide in slide_profiles:
            for keyword in slide.keywords:
                keyword_freq[keyword] = keyword_freq.get(keyword, 0) + 1
        
        return dict(sorted(keyword_freq.items(), key=lambda x: x[1], reverse=True)[:20])


class NarrativeSlideSelector:
    """Select slides based on learned organizational patterns"""
    
    def __init__(self, groq_api_key: str):
        self.groq_client = Groq(api_key=groq_api_key)
        self.profiler = OrgNarrativeProfiler(self.groq_client)
        self.org_profile = None
        self.slides = None
    
    def initialize(self, xml_path: Path, org_name: str, force_retrain: bool = False):
        """Initialize selector with profile (load or create)"""
        
        # Check for existing profile
        if not force_retrain and self.profiler.profile_exists(org_name):
            self.org_profile = self.profiler.load_profile(org_name)
        else:
            # Build new profile
            self.org_profile = self.profiler.build_profile(xml_path, org_name)
            self.profiler.save_profile(org_name, self.org_profile)
        
        # Load all slides for scoring
        self.slides = self.profiler._load_slides(xml_path)
        
        # Restore analysis data
        if 'slide_profiles' in self.org_profile:
            for i, profile_data in enumerate(self.org_profile['slide_profiles']):
                if i < len(self.slides):
                    self.slides[i].narrative_analysis = profile_data.narrative_analysis
                    self.slides[i].keywords = profile_data.keywords
    
    def select_slides(self, query: str, top_k: int = 5) -> List[Tuple[SlideProfile, float, Dict]]:
        """Select slides by narrative fit"""
        
        if not self.org_profile or not self.slides:
            print("Error: Not initialized. Call initialize() first.")
            return []
        
        print(f"\n{'='*80}")
        print(f"NARRATIVE-DRIVEN SLIDE SELECTION")
        print(f"{'='*80}\n")
        print(f"Query: {query}\n")
        
        scored_slides = []
        
        for slide in self.slides:
            # Score query match
            query_score = self._score_query_match(slide, query)
            
            # Score narrative alignment
            narrative_score = self._score_narrative_alignment(slide)
            
            # Score pattern fit
            pattern_score = self._score_pattern_fit(slide)
            
            # Combined score
            combined = (
                query_score * 0.50 +
                narrative_score * 0.30 +
                pattern_score * 0.20
            )
            
            breakdown = {
                'query_match': query_score,
                'narrative_alignment': narrative_score,
                'pattern_fit': pattern_score
            }
            
            scored_slides.append((slide, combined, breakdown))
        
        # Sort by score
        scored_slides.sort(key=lambda x: x[1], reverse=True)
        
        # Display results
        print(f"TOP {top_k} SLIDES:\n")
        
        for rank, (slide, score, breakdown) in enumerate(scored_slides[:top_k], 1):
            print(f"RANK {rank}: Slide #{slide.slide_index} (ID: {slide.slide_id})")
            print(f"  Score: {score:.4f}")
            print(f"  Query Match: {breakdown['query_match']:.3f}")
            print(f"  Narrative Fit: {breakdown['narrative_alignment']:.3f}")
            print(f"  Pattern Fit: {breakdown['pattern_fit']:.3f}")
            print(f"  Story Type: {slide.narrative_analysis.get('story_type', 'N/A')}")
            print(f"  Tone: {slide.narrative_analysis.get('tone', 'N/A')}")
            print(f"  Content: {slide.raw_text[:100]}...")
            print()
        
        return scored_slides[:top_k]
    
    def _score_query_match(self, slide: SlideProfile, query: str) -> float:
        """Score how well slide content matches query"""
        
        query_lower = query.lower()
        slide_text = slide.raw_text.lower()
        
        # Direct text match
        if query_lower in slide_text:
            return 1.0
        
        # Keyword match
        query_words = set(query_lower.split())
        slide_words = set(slide_text.split())
        overlap = len(query_words & slide_words)
        
        return min(overlap / len(query_words), 1.0) if query_words else 0.5
    
    def _score_narrative_alignment(self, slide: SlideProfile) -> float:
        """Score alignment with org narrative patterns"""
        
        if not self.org_profile.get('patterns'):
            return 0.5
        
        analysis = slide.narrative_analysis
        slide_tone = analysis.get('tone', 'professional')
        slide_flow = analysis.get('flow', 'linear')
        
        for pattern in self.org_profile['patterns'][:3]:
            if pattern.tone == slide_tone and pattern.flow_structure == slide_flow:
                return min(pattern.frequency_in_org + 0.3, 1.0)
        
        return 0.5
    
    def _score_pattern_fit(self, slide: SlideProfile) -> float:
        """Score how well slide fits org patterns"""
        
        analysis = slide.narrative_analysis
        narrative_strength = analysis.get('narrative_strength', 0.5)
        
        return float(narrative_strength)


def main():
    """Main execution"""
    import sys
    
    api_key = os.getenv('GROQ_API_KEY')
    if not api_key:
        print("Error: GROQ_API_KEY not set")
        sys.exit(1)
    
    if len(sys.argv) < 3:
        print('Usage: python solution.py <slides.xml> "Your query" [--force-retrain]')
        print('Example: python solution.py slides.xml "Q4 financial results"')
        sys.exit(1)
    
    xml_path = Path(sys.argv[1])
    query = sys.argv[2]
    force_retrain = '--force-retrain' in sys.argv
    
    if not xml_path.exists():
        print(f"Error: {xml_path} not found")
        sys.exit(1)
    
    # Extract org name from file
    org_name = xml_path.stem.split('_')[0][:20]
    
    # Initialize selector
    selector = NarrativeSlideSelector(api_key)
    
    print(f"\n{'='*80}")
    print(f"ORGANIZATION NARRATIVE-DRIVEN SLIDE SELECTOR")
    print(f"{'='*80}")
    print(f"\nOrganization: {org_name}")
    print(f"Data file: {xml_path.name}")
    
    # Initialize (load profile or create new)
    selector.initialize(xml_path, org_name, force_retrain=force_retrain)
    
    print(f"\n{'='*80}")
    print(f"ORG PROFILE SUMMARY")
    print(f"{'='*80}\n")
    print(f"Primary Patterns:")
    for pattern in selector.org_profile.get('patterns', [])[:3]:
        print(f"  • {pattern.pattern_name} ({pattern.frequency_in_org*100:.1f}%)")
    
    print(f"\nStorytelling Values:")
    for key, val in sorted(selector.org_profile.get('storytelling_values', {}).items(),
                          key=lambda x: x[1], reverse=True)[:3]:
        print(f"  • {key}: {val*100:.1f}%")
    
    print(f"\nTop Keywords: {', '.join(list(selector.org_profile.get('keywords_frequency', {}).keys())[:5])}\n")
    
    # Select slides
    results = selector.select_slides(query, top_k=5)
    
    if results:
        winner = results[0]
        print(f"\n{'='*80}")
        print(f"🏆 BEST MATCH: Slide #{winner[0].slide_index}")
        print(f"{'='*80}")


if __name__ == "__main__":
    main()