"""
Command Line Interface for PDF to Word Converter
"""

import click
import sys
from pathlib import Path
from tqdm import tqdm
from colorama import init, Fore, Style
import os

# Initialize colorama
init()

from core.converter import PDFToWordConverter
from utils.helpers import validate_file, get_file_size

@click.group()
@click.version_option(version="1.0.0")
def main():
    """PDF to Word Converter - Convert PDF files to Word documents with high fidelity."""
    pass

@main.command()
@click.argument('input_file', type=click.Path(exists=True, dir_okay=False))
@click.argument('output_file', type=click.Path(dir_okay=False))
@click.option('--quality', '-q', type=click.Choice(['low', 'medium', 'high']), 
              default='high', help='Image quality for conversion')
@click.option('--preserve-layout', '-p', is_flag=True, default=True,
              help='Preserve original layout (default: True)')
@click.option('--extract-images', '-i', is_flag=True, default=True,
              help='Extract and include images (default: True)')
@click.option('--extract-tables', '-t', is_flag=True, default=True,
              help='Extract and include tables (default: True)')
@click.option('--verbose', '-v', is_flag=True, help='Enable verbose output')
def convert(input_file, output_file, quality, preserve_layout, extract_images, extract_tables, verbose):
    """Convert a PDF file to Word document."""
    
    # Print header
    print(f"\n{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
    print(f"{Fore.CYAN}  PDF to Word Converter v1.0.0{Style.RESET_ALL}")
    print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}\n")
    
    try:
        # Validate input file
        input_path = Path(input_file)
        output_path = Path(output_file)
        
        if not validate_file(input_path, '.pdf'):
            click.echo(f"{Fore.RED}Error: Invalid PDF file.{Style.RESET_ALL}", err=True)
            sys.exit(1)
        
        # Ensure output directory exists
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Display file info
        file_size = get_file_size(input_path)
        print(f"{Fore.GREEN}📄 Input file:{Style.RESET_ALL} {input_path.name}")
        print(f"{Fore.GREEN}📊 File size:{Style.RESET_ALL} {file_size}")
        print(f"{Fore.GREEN}📝 Output file:{Style.RESET_ALL} {output_path.name}")
        print(f"{Fore.GREEN}⚙️  Quality:{Style.RESET_ALL} {quality}")
        print()
        
        # Configuration
        config = {
            'quality': quality,
            'preserve_layout': preserve_layout,
            'extract_images': extract_images,
            'extract_tables': extract_tables,
            'verbose': verbose
        }
        
        # Initialize converter
        converter = PDFToWordConverter(config)
        
        # Convert with progress bar
        with tqdm(total=100, desc="Converting", ncols=70, 
                 bar_format='{l_bar}{bar}| {percentage:3.0f}% {desc}') as pbar:
            
            def progress_callback(current, total, message=""):
                percentage = int((current / total) * 100)
                pbar.n = percentage
                pbar.set_description(f"Converting - {message}")
                pbar.refresh()
            
            # Perform conversion
            success = converter.convert(input_path, output_path, progress_callback)
            
            # Complete progress bar
            pbar.n = 100
            pbar.set_description("Conversion completed")
            pbar.refresh()
        
        if success:
            print(f"\n{Fore.GREEN}✅ Conversion successful!{Style.RESET_ALL}")
            print(f"{Fore.GREEN}📄 Output saved to:{Style.RESET_ALL} {output_path.absolute()}")
            
            # Show statistics if verbose
            if verbose:
                stats = converter.get_stats()
                print(f"\n{Fore.CYAN}📊 Conversion Statistics:{Style.RESET_ALL}")
                print(f"   Pages processed: {stats.get('pages', 0)}")
                print(f"   Images extracted: {stats.get('images', 0)}")
                print(f"   Tables extracted: {stats.get('tables', 0)}")
                print(f"   Processing time: {stats.get('time', 0):.2f}s")
        else:
            print(f"\n{Fore.RED}❌ Conversion failed!{Style.RESET_ALL}")
            sys.exit(1)
    
    except Exception as e:
        print(f"\n{Fore.RED}❌ Error: {str(e)}{Style.RESET_ALL}")
        if verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)

@main.command()
@click.argument('input_dir', type=click.Path(exists=True, file_okay=False))
@click.argument('output_dir', type=click.Path(file_okay=False))
@click.option('--quality', '-q', type=click.Choice(['low', 'medium', 'high']), 
              default='high', help='Image quality for conversion')
@click.option('--preserve-layout', '-p', is_flag=True, default=True,
              help='Preserve original layout (default: True)')
@click.option('--extract-images', '-i', is_flag=True, default=True,
              help='Extract and include images (default: True)')
@click.option('--extract-tables', '-t', is_flag=True, default=True,
              help='Extract and include tables (default: True)')
@click.option('--verbose', '-v', is_flag=True, help='Enable verbose output')
def batch(input_dir, output_dir, quality, preserve_layout, extract_images, extract_tables, verbose):
    """Convert multiple PDF files in a directory to Word documents."""
    
    # Print header
    print(f"\n{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
    print(f"{Fore.CYAN}  PDF to Word Converter - Batch Mode{Style.RESET_ALL}")
    print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}\n")
    
    try:
        input_path = Path(input_dir)
        output_path = Path(output_dir)
        
        # Find all PDF files
        pdf_files = list(input_path.glob("*.pdf"))
        
        if not pdf_files:
            print(f"{Fore.YELLOW}⚠️  No PDF files found in {input_dir}{Style.RESET_ALL}")
            return
        
        print(f"{Fore.GREEN}📁 Found {len(pdf_files)} PDF files{Style.RESET_ALL}")
        print(f"{Fore.GREEN}📁 Output directory:{Style.RESET_ALL} {output_path.absolute()}")
        print()
        
        # Create output directory
        output_path.mkdir(parents=True, exist_ok=True)
        
        # Configuration
        config = {
            'quality': quality,
            'preserve_layout': preserve_layout,
            'extract_images': extract_images,
            'extract_tables': extract_tables,
            'verbose': verbose
        }
        
        # Initialize converter
        converter = PDFToWordConverter(config)
        
        # Process each file
        successful = 0
        failed = 0
        
        for i, pdf_file in enumerate(pdf_files, 1):
            output_file = output_path / f"{pdf_file.stem}.docx"
            
            print(f"{Fore.CYAN}[{i}/{len(pdf_files)}]{Style.RESET_ALL} Processing: {pdf_file.name}")
            
            try:
                with tqdm(total=100, desc="Converting", ncols=50, leave=False) as pbar:
                    def progress_callback(current, total, message=""):
                        percentage = int((current / total) * 100)
                        pbar.n = percentage
                        pbar.refresh()
                    
                    success = converter.convert(pdf_file, output_file, progress_callback)
                    
                    if success:
                        print(f"  {Fore.GREEN}✅ Success{Style.RESET_ALL}")
                        successful += 1
                    else:
                        print(f"  {Fore.RED}❌ Failed{Style.RESET_ALL}")
                        failed += 1
                        
            except Exception as e:
                print(f"  {Fore.RED}❌ Error: {str(e)}{Style.RESET_ALL}")
                failed += 1
        
        # Summary
        print(f"\n{Fore.CYAN}📊 Batch Conversion Summary:{Style.RESET_ALL}")
        print(f"   Total files: {len(pdf_files)}")
        print(f"   Successful: {Fore.GREEN}{successful}{Style.RESET_ALL}")
        print(f"   Failed: {Fore.RED}{failed}{Style.RESET_ALL}")
        
    except Exception as e:
        print(f"\n{Fore.RED}❌ Error: {str(e)}{Style.RESET_ALL}")
        if verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()