"""
UI Variant Launcher
Run this script to choose and launch different UI variants.
"""

import os
import sys
import subprocess

def main():
    print("\n" + "="*50)
    print("   VOICE CONTROL - UI VARIANTS")
    print("="*50)
    print("\nSelect a UI variant to launch:\n")
    print("  1. ttkbootstrap  - Bootstrap-inspired modern theme")
    print("  2. sv-ttk        - Windows 11 native look")
    print("  3. customtkinter - Most modern/sleek (sidebar layout)")
    print("  4. Exit")
    print()
    
    choice = input("Enter your choice (1-4): ").strip()
    
    variants = {
        "1": "ttkbootstrap",
        "2": "sv_ttk",
        "3": "customtkinter"
    }
    
    if choice in variants:
        variant = variants[choice]
        variant_path = os.path.join(os.path.dirname(__file__), variant, "main.py")
        
        if os.path.exists(variant_path):
            print(f"\nLaunching {variant}...")
            os.chdir(os.path.join(os.path.dirname(__file__), variant))
            subprocess.run([sys.executable, "main.py"])
        else:
            print(f"Error: {variant_path} not found!")
    elif choice == "4":
        print("Goodbye!")
    else:
        print("Invalid choice!")

if __name__ == "__main__":
    main()
