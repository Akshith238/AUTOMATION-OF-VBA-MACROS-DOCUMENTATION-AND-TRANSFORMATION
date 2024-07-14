import google.generativeai as genai
import logging

logger = logging.getLogger(__name__)

# Configure the Gemini API
genai.configure(api_key='')

def enhance_explanation_with_gemini(explanation):
    try:
        logger.info(f"Enhancing explanation: {explanation}")
        model = genai.GenerativeModel('gemini-pro')
        prompt = f"Enhance and expand on this VBA macro explanation. Provide a detailed explanation while maintaining the structure (Name, Type, Purpose, Inputs, Process, Outputs, Business Impact):\n\n{explanation}\n\n"
        
        response = model.generate_content(prompt)
        
        enhanced = response.text
        logger.info(f"Enhanced explanation: {enhanced}")
        return enhanced
    except Exception as e:
        logger.error(f"Error enhancing explanation: {str(e)}", exc_info=True)
        return f"Error enhancing explanation: {str(e)}"
