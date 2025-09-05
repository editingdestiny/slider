#!/usr/bin/env python3
"""
Test script to generate PowerPoint presentation using slider.py and sample_esg_data.json
"""

import sys
import os
import json
from datetime import datetime

# Add the python_service directory to the path so we can import slider
sys.path.append('/home/sd22750/slider/python_service')

# Set matplotlib backend to Agg to avoid display issues
import matplotlib
matplotlib.use('Agg')

try:
    # Import the ESG_Presentation class from slider.py
    from slider import ESG_Presentation
    print("✅ Successfully imported ESG_Presentation from slider.py")
except ImportError as e:
    print(f"❌ Error importing from slider.py: {e}")
    sys.exit(1)

def test_esg_presentation():
    """Test the ESG presentation generation"""
    
    # Load the sample JSON data
    json_file = '/home/sd22750/slider/sample_esg_data.json'
    try:
        with open(json_file, 'r') as f:
            esg_data = json.load(f)
        print(f"✅ Successfully loaded data from {json_file}")
        print(f"   - Data keys: {list(esg_data.keys())}")
    except FileNotFoundError:
        print(f"❌ JSON file not found: {json_file}")
        return False
    except json.JSONDecodeError as e:
        print(f"❌ Error parsing JSON: {e}")
        return False
    
    try:
        # Create the ESG presentation
        print("\n🔄 Creating ESG presentation...")
        presentation = ESG_Presentation(esg_data)
        print("✅ ESG_Presentation object created successfully")
        
        # Add title slide
        print("🔄 Adding title slide...")
        presentation.add_title_slide(
            "ESG Analysis Report", 
            f"Comprehensive ESG Trends Analysis - {datetime.now().strftime('%B %Y')}"
        )
        print("✅ Title slide added")
        
        # Add summary slide with charts
        print("🔄 Adding executive summary slide...")
        presentation.add_slide1_summary()
        print("✅ Executive summary slide added")
        
        # Add business impact analysis slides
        print("🔄 Adding business impact analysis slides...")
        presentation.add_paginated_impact_slide()
        print("✅ Business impact analysis slides added")
        
        # Add regional trends slides
        print("🔄 Adding regional trends slides...")
        presentation.add_paginated_regional_trends()
        print("✅ Regional trends slides added")
        
        # Add sentiment justification slides
        print("🔄 Adding sentiment justification slides...")
        presentation.add_sentiment_justification_slides()
        print("✅ Sentiment justification slides added")
        
        # Add sources slides
        print("🔄 Adding data sources slides...")
        presentation.add_paginated_sources()
        print("✅ Data sources slides added")
        
        # Save the presentation
        output_file = '/home/sd22750/slider/test_esg_presentation.pptx'
        print(f"\n🔄 Saving presentation to {output_file}...")
        presentation.prs.save(output_file)
        print(f"✅ Presentation saved successfully!")
        
        # Get file size for verification
        file_size = os.path.getsize(output_file)
        print(f"   - File size: {file_size:,} bytes")
        print(f"   - Number of slides: {len(presentation.prs.slides)}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error creating presentation: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("🎯 Testing ESG PowerPoint Generation")
    print("=" * 50)
    
    success = test_esg_presentation()
    
    print("\n" + "=" * 50)
    if success:
        print("🎉 Test completed successfully!")
        print("📄 PowerPoint file: /home/sd22750/slider/test_esg_presentation.pptx")
    else:
        print("💥 Test failed!")
        sys.exit(1)
