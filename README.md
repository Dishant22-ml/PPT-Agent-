PowerPoint AI Pipeline 🎨🤖
A comprehensive Python-based system for extracting, analyzing, selecting, modifying, and regenerating PowerPoint presentations using AI and machine learning techniques.
📋 Table of Contents

Overview
Architecture
Features
Installation
Quick Start
Detailed Usage
Project Components
Examples
Troubleshooting
Contributing


🎯 Overview
This project provides a complete pipeline for working with PowerPoint presentations programmatically. It combines traditional computational analysis with cutting-edge Large Language Model (LLM) capabilities to understand, select, and modify slides intelligently.
What Can This Do?

Extract all features from PPTX files into structured XML format
Analyze slide content using hybrid mathematical + AI approaches
Select relevant slides based on organizational storytelling patterns
Modify slide content intelligently using natural language prompts
Generate new PPTX files with your modifications applied

Use Cases

📊 Content Repurposing: Transform presentations for different audiences
🔍 Smart Search: Find relevant slides across large presentation libraries
✏️ Automated Updates: Batch update presentations with new information
🎭 Style Transfer: Apply organizational narrative patterns to new content
🤖 AI-Powered Editing: Modify slides using natural language commands


🏗️ Architecture
The system consists of four main components that work together:
┌─────────────────┐
│  1. EXTRACTOR   │  PowerPoint → XML
│  (slide_        │  Deep feature extraction
│   extractor.py) │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  2. SELECTOR    │  XML → Analysis → Selected Slides
│  (slide_        │  AI-powered narrative matching
│   selection.py) │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  3. MODIFIER    │  XML + Prompt → Modified XML
│  (slide_        │  Hybrid AI analysis & editing
│   modifier.py)  │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  4. GENERATOR   │  Original PPTX + Modified XML → New PPTX
│  (ppt_          │  Reverse engineering to PowerPoint
│   generator.py) │
└─────────────────┘

✨ Features
Component 1: Slide Extractor

Extracts all presentation features into structured XML
Captures text, images, charts, tables, shapes, and groups
Preserves formatting (fonts, colors, sizes, positions)
Extracts theme definitions and color palettes
Computes slide-level and presentation-level statistics
Converts colors to multiple formats (RGB, LAB, HEX)

Component 2: Slide Selector

Uses Groq LLM to analyze narrative patterns
Creates organizational "storytelling profiles"
Persistent storage - trains once, use forever
Scores slides based on:

Query match
Narrative alignment
Pattern fit


Identifies storytelling characteristics (tone, flow, style)

Component 3: Slide Modifier

Hybrid analysis: Mathematics + LLM reasoning
Understands slide structure semantically
Modifies content using natural language prompts
Shows before/after comparisons
Handles complex nested text structures
High confidence scoring for changes

Component 4: PPT Generator

Converts modified XML back to PowerPoint
Preserves original formatting
Updates text, colors, and positions
Maintains slide master relationships
Produces production-ready PPTX files


📦 Installation
Prerequisites

Python 3.8+ (recommended: 3.10 or 3.11)
Groq API Key (free at console.groq.com)

Step 1: Clone the Repository
bashgit clone https://github.com/yourusername/powerpoint-ai-pipeline.git
cd powerpoint-ai-pipeline
Step 2: Install Dependencies
bashpip install -r requirements.txt
Required packages:
txtpython-pptx==0.6.21
groq==0.4.2
lxml==4.9.3
Step 3: Set Up API Key
Option A: Environment Variable (Recommended)
bashexport GROQ_API_KEY="your_api_key_here"
Option B: Pass directly in commands
bashpython slide_modifier.py slide.xml your_api_key "prompt"
Step 4: Create Project Folders
bashmkdir -p data output org_profiles
```

**Folder structure:**
```
powerpoint-ai-pipeline/
├── data/              # Input PPTX files
├── output/            # Generated XML files
├── org_profiles/      # Saved organizational profiles
├── slide_extractor.py
├── slide_selection.py
├── slide_modifier.py
├── ppt_generator.py
└── README.md

🚀 Quick Start
Complete Workflow Example
bash# 1. Extract features from PowerPoint
python slide_extractor.py data/presentation.pptx output/

# 2. Select relevant slides (creates profile on first run)
python slide_selection.py output/presentation_training.xml "quarterly financial results"

# 3. Modify slide content
python slide_modifier.py output/presentation_training.xml $GROQ_API_KEY "Change all references from 'Graphite' to 'AI in Mining'"

# 4. Generate new PowerPoint
python ppt_generator.py data/presentation.pptx output/presentation_training_modified.xml output/modified_presentation.pptx

📚 Detailed Usage
Component 1: Slide Extractor
Purpose: Convert PowerPoint to machine-readable XML format
Command:
bashpython slide_extractor.py <input.pptx> [output_folder]
Examples:
bash# Basic extraction
python slide_extractor.py data/quarterly_report.pptx

# Specify output folder
python slide_extractor.py data/quarterly_report.pptx output/

# Extract from absolute path
python slide_extractor.py /home/user/presentations/deck.pptx
Output:

Creates <filename>_training.xml in output folder
XML contains:

Document metadata (author, dates, dimensions)
Theme definitions (colors, fonts, effects)
Slide masters and layouts
Complete slide content with:

Text with full formatting
Geometry (positions, sizes, rotations)
Images and charts
Tables and groups


Computed features (readability, balance, color diversity)
Global statistics



Key Features Extracted:

✅ Text content with character-level formatting
✅ Font families, sizes, colors, styles
✅ Paragraph alignment, spacing, bullets
✅ Shapes, images, charts, tables
✅ Positions normalized to 0-1 range
✅ Color palettes in RGB, LAB, HEX
✅ Semantic roles (title, body, data viz)


Component 2: Slide Selector
Purpose: Find relevant slides using AI-powered narrative analysis
Command:
bashpython slide_selection.py <slides.xml> "your search query" [--force-retrain]
Examples:
bash# First run - creates organizational profile
python slide_selection.py output/company_deck_training.xml "product launch strategy"

# Subsequent runs - uses cached profile (instant)
python slide_selection.py output/company_deck_training.xml "market expansion plans"

# Force rebuild profile with new data
python slide_selection.py output/updated_deck_training.xml "Q4 results" --force-retrain
```

**How It Works:**

1. **Profile Creation (First Run)**:
   - Analyzes ALL slides with LLM
   - Extracts narrative patterns
   - Identifies storytelling characteristics
   - Calculates organizational values
   - **Saves profile to disk** (`org_profiles/<name>_profile.pkl`)

2. **Selection (All Runs)**:
   - Loads profile (instant if cached)
   - Scores slides based on:
     - **Query Match** (50%): Direct content relevance
     - **Narrative Alignment** (30%): Fits org patterns
     - **Pattern Fit** (20%): Storytelling strength
   - Returns top 5 ranked slides

**Output:**
```
TOP 5 SLIDES:

RANK 1: Slide #12 (ID: slide12)
  Score: 0.8742
  Query Match: 0.950
  Narrative Fit: 0.780
  Pattern Fit: 0.650
  Story Type: problem-solution
  Tone: professional
  Content: Our market expansion strategy focuses on...

[... more results ...]
Profile Persistence:

✅ Train once per organization
✅ Instant searches thereafter
✅ No re-training needed for new queries
✅ Update profile only when presentations change significantly


Component 3: Slide Modifier
Purpose: Intelligently modify slide content using hybrid AI analysis
Command:
bashpython slide_modifier.py <slide.xml> <groq_api_key> [modification_prompt]
Examples:
bash# Analyze only (no modifications)
python slide_modifier.py output/slide1_training.xml $GROQ_API_KEY

# Modify title
python slide_modifier.py output/slide1_training.xml $GROQ_API_KEY "Change title to 'AI Revolution 2024'"

# Complex modification
python slide_modifier.py output/slide1_training.xml $GROQ_API_KEY "Replace all content about Graphite mining with AI in resource management, maintaining the same structure"

# Batch update multiple topics
python slide_modifier.py output/slide1_training.xml $GROQ_API_KEY "Update: Q3 → Q4, 2023 → 2024, 15% growth → 23% growth"
Analysis Process:

Phase 1: Mathematical Analysis

Position categorization (top, mid, bottom, left, center, right)
Size categorization (XS, S, M, L, XL)
Text statistics (length, word count)
Confidence scoring


Phase 2: LLM Semantic Analysis

Understands element purpose (title, subtitle, body, image, chart)
Assigns semantic roles
Reasons about content meaning
Provides confidence scores


Phase 3: Consensus Fusion

Combines mathematical + AI insights
Creates unified understanding
Produces final categorization



Modification Features:

✅ Comprehensive text replacement
✅ Maintains formatting and structure
✅ Handles nested content (bullets, sub-points)
✅ Shows before/after comparisons
✅ Multiple fallback strategies
✅ Smart element matching

Output:

Creates <filename>_modified.xml
Displays all modifications with confidence scores
Shows old → new values for verification
Saves analysis to hybrid_analysis.json


Component 4: PPT Generator
Purpose: Convert modified XML back to PowerPoint format
Command:
bashpython ppt_generator.py <original.pptx> <modified.xml> <output.pptx>
Examples:
bash# Generate new presentation
python ppt_generator.py data/template.pptx output/presentation_modified.xml output/final_presentation.pptx

# Use different template
python ppt_generator.py data/branded_template.pptx output/content_modified.xml output/branded_output.pptx
Process:

Loads original PPTX as template (preserves master slides)
Parses modified XML
Finds corresponding shapes by ID
Updates text content and formatting
Saves new PPTX file

What Gets Updated:

✅ Text content
✅ Font properties (family, size, color)
✅ Paragraph formatting (alignment, spacing, bullets)
✅ Text styles (bold, italic, underline)
✅ Alt text (accessibility)

What Gets Preserved:

✅ Slide layouts
✅ Master slides
✅ Images (unless explicitly replaced)
✅ Charts and tables structure
✅ Animations and transitions
✅ Theme and branding


💡 Examples
Example 1: Extract and Analyze
bash# Extract features
python slide_extractor.py data/company_overview.pptx output/

# Analyze a specific slide
python slide_modifier.py output/company_overview_training.xml $GROQ_API_KEY

# View analysis
cat hybrid_analysis.json
Example 2: Content Repurposing
bash# Extract from original deck
python slide_extractor.py data/internal_deck.pptx output/

# Modify for external audience
python slide_modifier.py output/internal_deck_training.xml $GROQ_API_KEY "Remove all confidential information and replace with public-facing content"

# Generate client version
python ppt_generator.py data/internal_deck.pptx output/internal_deck_training_modified.xml output/client_deck.pptx
Example 3: Smart Slide Library
bash# Build profile from company presentations
python slide_selection.py output/all_presentations_training.xml "dummy query" --force-retrain

# Search for specific content
python slide_selection.py output/all_presentations_training.xml "customer testimonials"
python slide_selection.py output/all_presentations_training.xml "pricing models"
python slide_selection.py output/all_presentations_training.xml "competitive analysis"
Example 4: Batch Updates
bash# Extract slides
python slide_extractor.py data/q3_report.pptx output/

# Update all Q3 → Q4 references
python slide_modifier.py output/q3_report_training.xml $GROQ_API_KEY "Update all references: Q3 2024 → Q4 2024, September → December, $2.5M revenue → $3.1M revenue"

# Generate updated deck
python ppt_generator.py data/q3_report.pptx output/q3_report_training_modified.xml output/q4_report.pptx

🐛 Troubleshooting
Common Issues
1. "GROQ_API_KEY not set"
bash# Solution: Export your API key
export GROQ_API_KEY="gsk_your_key_here"

# Or pass directly
python slide_modifier.py slide.xml gsk_your_key_here "prompt"
2. "File not found" error
bash# Check file exists
ls -la data/presentation.pptx

# Use absolute path
python slide_extractor.py /full/path/to/presentation.pptx
3. "Element ID not found" during modification

This happens when LLM suggests non-existent IDs
The modifier has fallback strategies
Check hybrid_analysis.json for available element IDs
Some modifications may fail but others will succeed

4. "Could not parse JSON response" from LLM

Network issue or API rate limit
Script falls back to mathematical analysis only
Retry the command after a few seconds

5. Text not updating in generated PPTX

Original PPTX may have text in master slides
Try unlocking/ungrouping shapes in PowerPoint first
Check if shape IDs match between XML and PPTX

Performance Tips

Extractor: Large files (>50MB) may take 2-3 minutes
Selector: First run trains profile (2-5 min), subsequent runs instant
Modifier: LLM calls take 5-15 seconds depending on content
Generator: Fast (<10 seconds for most files)

Getting Help

Check verbose output for detailed error messages
Review hybrid_analysis.json for element details
Open an issue on GitHub with:

Command you ran
Error message
XML file (if possible)




🤝 Contributing
Contributions welcome! Areas for improvement:

 Support for more LLM providers (OpenAI, Anthropic, local models)
 Image generation/replacement with DALL-E or Stable Diffusion
 Chart data modification
 Animation preservation in generator
 Web interface for non-technical users
 Batch processing CLI
 PowerPoint add-in

To contribute:

Fork the repository
Create a feature branch (git checkout -b feature/amazing-feature)
Commit changes (git commit -m 'Add amazing feature')
Push to branch (git push origin feature/amazing-feature)
Open a Pull Request


📄 License
MIT License - feel free to use this in your projects!

🙏 Acknowledgments

python-pptx for PowerPoint manipulation
Groq for fast LLM inference
OpenXML specification for PPTX format insights


📞 Contact
Questions? Reach out:

GitHub Issues: Create an issue
Email: your.email@example.com
