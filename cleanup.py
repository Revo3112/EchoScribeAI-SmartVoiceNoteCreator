#!/usr/bin/env python3
"""
EchoScribe AI Project Cleanup Script
====================================

This script cleans up the project directory by removing:
- Python cache files (__pycache__ directories and .pyc files)
- Log files (*.log)
- Temporary files and directories
- Development/demo files

Usage:
    python cleanup.py

Author: EchoScribe AI Team
"""

import os
import shutil
import glob
from pathlib import Path

def clean_pycache():
    """Remove all __pycache__ directories and .pyc files."""
    print("🧹 Cleaning Python cache files...")

    # Remove __pycache__ directories
    for pycache_dir in Path('.').rglob('__pycache__'):
        if pycache_dir.is_dir():
            print(f"  Removing: {pycache_dir}")
            shutil.rmtree(pycache_dir)

    # Remove .pyc files
    for pyc_file in Path('.').rglob('*.pyc'):
        if pyc_file.is_file():
            print(f"  Removing: {pyc_file}")
            pyc_file.unlink()

def clean_logs():
    """Remove log files."""
    print("📝 Cleaning log files...")

    for log_file in Path('.').rglob('*.log'):
        if log_file.is_file():
            print(f"  Removing: {log_file}")
            log_file.unlink()

def clean_temp():
    """Remove temporary directories and files."""
    print("🗂️  Cleaning temporary files...")

    temp_dirs = ['temp', 'tmp', '.pytest_cache']
    for temp_dir in temp_dirs:
        temp_path = Path(temp_dir)
        if temp_path.exists() and temp_path.is_dir():
            print(f"  Removing: {temp_path}")
            shutil.rmtree(temp_path)

def clean_audio_outputs():
    """Remove audio output files that might be left from testing."""
    print("🎵 Cleaning audio test files...")

    audio_extensions = ['*.wav', '*.mp3', '*.m4a', '*.aac']
    for pattern in audio_extensions:
        for audio_file in Path('.').rglob(pattern):
            if audio_file.is_file() and 'output' in str(audio_file).lower():
                print(f"  Removing: {audio_file}")
                audio_file.unlink()

def main():
    """Main cleanup function."""
    print("🚀 Starting EchoScribe AI project cleanup...")
    print("=" * 50)

    try:
        clean_pycache()
        clean_logs()
        clean_temp()
        clean_audio_outputs()

        print("=" * 50)
        print("✅ Project cleanup completed successfully!")
        print("\n📋 Summary:")
        print("  - Python cache files removed")
        print("  - Log files removed")
        print("  - Temporary files removed")
        print("  - Audio test outputs removed")
        print("\n🎯 Project is now ready for Git operations!")

    except Exception as e:
        print(f"❌ Error during cleanup: {e}")
        return 1

    return 0

if __name__ == "__main__":
    exit(main())
