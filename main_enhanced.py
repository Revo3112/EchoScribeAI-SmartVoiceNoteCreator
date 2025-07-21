# -*- coding: utf-8 -*-
"""
Main Module for EchoScribe AI - Enhanced Version
Complete integration from monolithic system with 100% functionality.
"""

import logging
import sys
from pathlib import Path

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('echoscribe.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)

def main():
    """Main entry point for EchoScribe AI Enhanced."""
    try:
        logger.info("Starting EchoScribe AI Enhanced...")

        # Import UI after logging setup
        from src.ui.main_window import EchoScribeUI

        # Create and run the application
        app = EchoScribeUI()
        app.run()

    except ImportError as e:
        logger.error(f"Import error: {e}")
        logger.error("Please ensure all dependencies are installed:")
        logger.error("pip install customtkinter soundcard pyaudio groq speech_recognition python-docx matplotlib numpy")
        sys.exit(1)

    except Exception as e:
        logger.error(f"Fatal error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
