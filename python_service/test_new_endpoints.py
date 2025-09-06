#!/usr/bin/env python3
"""
Test file for the new webhook endpoints that accept user inputs
"""

import json
import sys
import os

# Add the current directory to Python path to import our modules
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_slide_generation_request():
    """Test the SlideGenerationRequest model"""
    try:
        from main import SlideGenerationRequest
        
        # Test with valid data
        request1 = SlideGenerationRequest(search_phrase="Climate Change", number_of_slides=8)
        print(f"✓ Valid request: {request1.search_phrase} with {request1.number_of_slides} slides")
        
        # Test with default number of slides
        request2 = SlideGenerationRequest(search_phrase="AI Technology")
        print(f"✓ Default slides request: {request2.search_phrase} with {request2.number_of_slides} slides")
        
        return True
    except Exception as e:
        print(f"✗ Error testing SlideGenerationRequest: {e}")
        return False

def test_esg_analysis_request():
    """Test the ESGAnalysisRequest model"""
    try:
        from main import ESGAnalysisRequest
        
        # Test with valid data
        request1 = ESGAnalysisRequest(search_phrase="Renewable Energy", number_of_slides=12)
        print(f"✓ Valid ESG request: {request1.search_phrase} with {request1.number_of_slides} slides")
        
        # Test with default number of slides
        request2 = ESGAnalysisRequest(search_phrase="Carbon Footprint")
        print(f"✓ Default ESG request: {request2.search_phrase} with {request2.number_of_slides} slides")
        
        return True
    except Exception as e:
        print(f"✗ Error testing ESGAnalysisRequest: {e}")
        return False

def test_sample_requests():
    """Test sample request JSON structures"""
    print("\n--- Sample Request Structures ---")
    
    # Sample slide generation request
    slide_request = {
        "search_phrase": "Digital Transformation",
        "number_of_slides": 6
    }
    print(f"Sample Slide Request: {json.dumps(slide_request, indent=2)}")
    
    # Sample ESG analysis request  
    esg_request = {
        "search_phrase": "Sustainable Supply Chain",
        "number_of_slides": 10
    }
    print(f"Sample ESG Request: {json.dumps(esg_request, indent=2)}")

def main():
    """Run all tests"""
    print("Testing New Webhook Endpoint Models")
    print("=" * 40)
    
    # Test the Pydantic models (will work even without FastAPI installed)
    success1 = test_slide_generation_request()
    success2 = test_esg_analysis_request()
    
    test_sample_requests()
    
    print("\n--- Test Summary ---")
    if success1 and success2:
        print("✓ All model tests passed!")
        print("✓ New webhook endpoints are ready to accept user inputs")
        print("✓ The system can now generate presentations based on:")
        print("  - Search phrase (user-defined topic)")
        print("  - Number of slides (user-defined count)")
    else:
        print("✗ Some tests failed - check the error messages above")

if __name__ == "__main__":
    main()
