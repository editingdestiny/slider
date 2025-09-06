"""
LLM Research Service for ESG and General Presentations
Integrates with OpenAI to generate research-based presentation content
"""

import os
import json
import logging
from typing import Dict, Any, List
from openai import OpenAI
from datetime import datetime

logger = logging.getLogger(__name__)

class LLMResearchService:
    def __init__(self):
        # Initialize OpenAI client
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            logger.warning("OPENAI_API_KEY not found. Using fallback mock data.")
            self.client = None
        else:
            self.client = OpenAI(api_key=api_key)
    
    async def generate_esg_research(self, search_phrase: str, number_of_slides: int) -> Dict[str, Any]:
        """
        Generate ESG research data based on search phrase
        Returns structured JSON matching sample_esg_data.json format
        """
        if not self.client:
            logger.info("Using mock ESG data (no OpenAI key)")
            return self._generate_mock_esg_data(search_phrase, number_of_slides)
        
        try:
            logger.info(f"Generating ESG research for: {search_phrase}")
            
            prompt = self._create_esg_research_prompt(search_phrase, number_of_slides)
            
            response = await self._call_openai(prompt, "gpt-4")
            
            # Parse the JSON response
            esg_data = json.loads(response)
            
            # Validate and enhance the data
            esg_data = self._validate_esg_data(esg_data, search_phrase)
            
            logger.info(f"Successfully generated ESG research for {search_phrase}")
            return esg_data
            
        except Exception as e:
            logger.error(f"Error generating ESG research: {str(e)}")
            return self._generate_mock_esg_data(search_phrase, number_of_slides)
    
    async def generate_general_research(self, search_phrase: str, number_of_slides: int) -> List[Dict[str, Any]]:
        """
        Generate general presentation research based on search phrase
        Returns list of slide data
        """
        if not self.client:
            logger.info("Using mock general data (no OpenAI key)")
            return self._generate_mock_general_data(search_phrase, number_of_slides)
        
        try:
            logger.info(f"Generating general research for: {search_phrase}")
            
            prompt = self._create_general_research_prompt(search_phrase, number_of_slides)
            
            response = await self._call_openai(prompt, "gpt-4")
            
            # Parse the JSON response
            slides_data = json.loads(response)
            
            logger.info(f"Successfully generated {len(slides_data)} slides for {search_phrase}")
            return slides_data
            
        except Exception as e:
            logger.error(f"Error generating general research: {str(e)}")
            return self._generate_mock_general_data(search_phrase, number_of_slides)
    
    async def _call_openai(self, prompt: str, model: str = "gpt-4") -> str:
        """Make async call to OpenAI API"""
        try:
            response = self.client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": "You are a professional research analyst who creates comprehensive, accurate presentations based on current market data and trends."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=4000
            )
            return response.choices[0].message.content
        except Exception as e:
            logger.error(f"OpenAI API call failed: {str(e)}")
            raise
    
    def _create_esg_research_prompt(self, search_phrase: str, number_of_slides: int) -> str:
        """Create prompt for ESG research"""
        return f"""
        Research and analyze ESG (Environmental, Social, Governance) factors related to "{search_phrase}".
        
        Create a comprehensive ESG analysis in JSON format that matches this structure:
        
        {{
            "executiveSummary": "Detailed summary of ESG findings for {search_phrase}",
            "sentimentSummary": {{
                "Positive": percentage,
                "Neutral": percentage, 
                "Negative": percentage
            }},
            "regionalData": {{
                "North America": {{
                    "United States": {{
                        "Environmental": [
                            {{
                                "theme": "specific theme related to {search_phrase}",
                                "sentiment": "Positive/Neutral/Negative",
                                "articleCount": number,
                                "justification": "detailed explanation"
                            }}
                        ],
                        "Social": [...],
                        "Governance": [...]
                    }}
                }},
                "Europe": {{ ... }},
                "Asia": {{ ... }}
            }},
            "impactAnalysis": [
                {{
                    "country": "country name",
                    "theme": "theme related to {search_phrase}",
                    "impactArea": "Business area affected", 
                    "impactLevel": "High/Medium/Low",
                    "rationale": "explanation of impact"
                }}
            ],
            "dataSources": [
                {{
                    "source": "credible source name",
                    "reliabilityScore": "X/10",
                    "justification": "why this source is reliable",
                    "url": "https://example.com"
                }}
            ]
        }}
        
        Focus on current, accurate information about {search_phrase} ESG implications.
        Include at least {number_of_slides} key themes across the environmental, social, and governance categories.
        Ensure percentages in sentimentSummary add up to 100.
        Return only valid JSON.
        """
    
    def _create_general_research_prompt(self, search_phrase: str, number_of_slides: int) -> str:
        """Create prompt for general presentation research"""
        return f"""
        Research "{search_phrase}" and create {number_of_slides} presentation slides with comprehensive, current information.
        
        Return a JSON array of slide objects in this format:
        
        [
            {{
                "title": "Compelling slide title about {search_phrase}",
                "headline": "Key insight or main message",
                "content": "• Detailed bullet point 1\\n• Research-backed point 2\\n• Strategic insight 3\\n• Actionable conclusion"
            }}
        ]
        
        Requirements:
        - Create exactly {number_of_slides} slides
        - Each slide should cover different aspects of {search_phrase}
        - Use current, factual information
        - Include strategic insights and implications
        - Format content as bullet points with \\n separators
        - Make titles specific and engaging
        - Ensure headlines capture key insights
        
        Research areas to cover (adapt to {search_phrase}):
        - Overview and current state
        - Market trends and opportunities  
        - Key challenges and solutions
        - Technology and innovation aspects
        - Strategic recommendations
        - Future outlook
        
        Return only valid JSON array.
        """
    
    def _validate_esg_data(self, data: Dict[str, Any], search_phrase: str) -> Dict[str, Any]:
        """Validate and enhance ESG data structure"""
        # Ensure required keys exist
        required_keys = ["executiveSummary", "sentimentSummary", "regionalData", "impactAnalysis", "dataSources"]
        for key in required_keys:
            if key not in data:
                logger.warning(f"Missing key {key} in ESG data, adding default")
                data[key] = self._get_default_esg_value(key, search_phrase)
        
        # Validate sentiment percentages
        if "sentimentSummary" in data:
            sentiment = data["sentimentSummary"]
            total = sentiment.get("Positive", 0) + sentiment.get("Neutral", 0) + sentiment.get("Negative", 0)
            if total != 100:
                logger.warning(f"Sentiment percentages don't add to 100 ({total}), normalizing")
                # Normalize to 100
                factor = 100 / total if total > 0 else 1
                sentiment["Positive"] = round(sentiment.get("Positive", 0) * factor)
                sentiment["Neutral"] = round(sentiment.get("Neutral", 0) * factor)  
                sentiment["Negative"] = 100 - sentiment["Positive"] - sentiment["Neutral"]
        
        return data
    
    def _get_default_esg_value(self, key: str, search_phrase: str) -> Any:
        """Get default values for missing ESG data keys"""
        defaults = {
            "executiveSummary": f"ESG analysis for {search_phrase} focusing on environmental, social, and governance factors.",
            "sentimentSummary": {"Positive": 60, "Neutral": 25, "Negative": 15},
            "regionalData": {"North America": {"United States": {"Environmental": [], "Social": [], "Governance": []}}},
            "impactAnalysis": [],
            "dataSources": [{"source": "Industry Research", "reliabilityScore": "7/10", "justification": "General industry analysis", "url": "https://example.com"}]
        }
        return defaults.get(key, {})
    
    def _generate_mock_esg_data(self, search_phrase: str, number_of_slides: int) -> Dict[str, Any]:
        """Generate mock ESG data when OpenAI is not available"""
        return {
            "executiveSummary": f"Mock ESG analysis for {search_phrase}. This analysis examines environmental, social, and governance factors related to {search_phrase}, highlighting key trends, opportunities, and risks in the current market landscape.",
            "sentimentSummary": {
                "Positive": 65,
                "Neutral": 20, 
                "Negative": 15
            },
            "regionalData": {
                "North America": {
                    "United States": {
                        "Environmental": [
                            {
                                "theme": f"{search_phrase} - Environmental Impact",
                                "sentiment": "Positive",
                                "articleCount": 12,
                                "justification": f"Strong environmental initiatives related to {search_phrase}"
                            }
                        ],
                        "Social": [
                            {
                                "theme": f"{search_phrase} - Social Responsibility",
                                "sentiment": "Positive", 
                                "articleCount": 8,
                                "justification": f"Positive social impact from {search_phrase} initiatives"
                            }
                        ],
                        "Governance": [
                            {
                                "theme": f"{search_phrase} - Corporate Governance",
                                "sentiment": "Neutral",
                                "articleCount": 6,
                                "justification": f"Mixed governance practices in {search_phrase} sector"
                            }
                        ]
                    }
                }
            },
            "impactAnalysis": [
                {
                    "country": "United States",
                    "theme": f"{search_phrase} - Strategic Impact",
                    "impactArea": "Risk Management",
                    "impactLevel": "Medium",
                    "rationale": f"Moderate impact expected from {search_phrase} on overall business strategy"
                }
            ],
            "dataSources": [
                {
                    "source": "Mock ESG Research Database",
                    "reliabilityScore": "8/10",
                    "justification": f"Comprehensive mock data for {search_phrase} analysis",
                    "url": "https://example.com/esg"
                }
            ]
        }
    
    def _generate_mock_general_data(self, search_phrase: str, number_of_slides: int) -> List[Dict[str, Any]]:
        """Generate mock general presentation data when OpenAI is not available"""
        slides = []
        
        slide_templates = [
            {
                "title": f"Introduction to {search_phrase}",
                "headline": f"Understanding the fundamentals of {search_phrase}",
                "content": f"• Overview of {search_phrase} landscape\n• Key market players and stakeholders\n• Current industry trends\n• Strategic importance"
            },
            {
                "title": f"{search_phrase} Market Analysis", 
                "headline": f"Current state and trends in {search_phrase}",
                "content": f"• Market size and growth projections\n• Competitive landscape analysis\n• Emerging opportunities\n• Key challenges and barriers"
            },
            {
                "title": f"Technology & Innovation in {search_phrase}",
                "headline": f"Technological advances shaping {search_phrase}",
                "content": f"• Latest technological developments\n• Innovation drivers and catalysts\n• Disruptive technologies on horizon\n• Impact on business models"
            },
            {
                "title": f"Strategic Implications of {search_phrase}",
                "headline": f"How {search_phrase} affects business strategy",
                "content": f"• Strategic opportunities for growth\n• Risk mitigation strategies\n• Investment considerations\n• Competitive advantages"
            },
            {
                "title": f"Implementation Roadmap for {search_phrase}",
                "headline": f"Practical steps for {search_phrase} adoption",
                "content": f"• Short-term implementation priorities\n• Medium-term strategic initiatives\n• Long-term vision and goals\n• Success metrics and KPIs"
            },
            {
                "title": f"Future Outlook for {search_phrase}",
                "headline": f"Predictions and trends for {search_phrase}",
                "content": f"• Future market projections\n• Emerging trends to watch\n• Potential disruptors\n• Strategic recommendations"
            }
        ]
        
        # Select slides up to the requested number
        for i in range(min(number_of_slides, len(slide_templates))):
            slides.append(slide_templates[i])
        
        # If more slides requested than templates, create additional ones
        while len(slides) < number_of_slides:
            slide_num = len(slides) + 1
            slides.append({
                "title": f"{search_phrase} - Additional Insights {slide_num}",
                "headline": f"Further analysis of {search_phrase}",
                "content": f"• Additional research findings\n• Supplementary market data\n• Extended strategic analysis\n• Continued recommendations"
            })
        
        return slides
