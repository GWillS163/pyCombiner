"""
Helper utility functions
"""

def greet(name: str) -> str:
    """
    Generate a greeting message
    
    Args:
        name: Name to greet
        
    Returns:
        Greeting message
    """
    return f"Hello, {name}!"

def format_output(text: str) -> str:
    return f"[PROCESSED] {text.upper()}" 