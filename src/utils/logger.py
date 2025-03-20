#!/usr/bin/env python3
"""
Logger module for the application.
"""

import logging
import sys

def setup_logger(name='pdf_generator', level=logging.INFO):
    """
    Set up and configure logger.
    
    Args:
        name: Logger name
        level: Logging level (default: INFO)
        
    Returns:
        logging.Logger: Configured logger instance
    """
    # Create logger
    logger = logging.getLogger(name)
    logger.setLevel(level)
    
    # Create console handler and set level
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    
    # Create formatter
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    
    # Add formatter to console handler
    console_handler.setFormatter(formatter)
    
    # Add console handler to logger
    logger.addHandler(console_handler)
    
    return logger 