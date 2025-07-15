"""
Funciones auxiliares para el convertidor PDF a Word
"""

import os
import sys
from pathlib import Path
from typing import List, Tuple, Optional
import fitz  # PyMuPDF
from colorama import Fore, Style


def validate_file(file_path: str, extension: str = '.pdf') -> bool:
    """
    Valida si un archivo existe y tiene la extensión correcta
    
    Args:
        file_path: Ruta al archivo
        extension: Extensión esperada (por defecto .pdf)
    
    Returns:
        bool: True si el archivo es válido
    """
    if not os.path.exists(file_path):
        print(f"{Fore.RED}❌ Error: El archivo no existe: {file_path}{Style.RESET_ALL}")
        return False
    
    if not file_path.lower().endswith(extension.lower()):
        print(f"{Fore.RED}❌ Error: El archivo debe tener extensión {extension}{Style.RESET_ALL}")
        return False
    
    return True


def get_file_size(file_path: str) -> str:
    """
    Obtiene el tamaño de un archivo en formato legible
    
    Args:
        file_path: Ruta al archivo
    
    Returns:
        str: Tamaño del archivo (ej: "2.5 MB")
    """
    try:
        size_bytes = os.path.getsize(file_path)
        
        # Convertir bytes a unidades legibles
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024.0
        
        return f"{size_bytes:.1f} TB"
    except OSError:
        return "Desconocido"


def get_pdf_info(pdf_path: str) -> dict:
    """
    Obtiene información básica de un archivo PDF
    
    Args:
        pdf_path: Ruta al archivo PDF
    
    Returns:
        dict: Información del PDF (páginas, tamaño, etc.)
    """
    try:
        doc = fitz.open(pdf_path)
        info = {
            'pages': len(doc),
            'size': get_file_size(pdf_path),
            'title': doc.metadata.get('title', 'Sin título'),
            'author': doc.metadata.get('author', 'Desconocido'),
            'has_text': False,
            'has_images': False
        }
        
        # Verificar si tiene texto seleccionable
        for page_num in range(min(3, len(doc))):  # Revisar solo primeras 3 páginas
            page = doc[page_num]
            text = page.get_text()
            if text.strip():
                info['has_text'] = True
                break
        
        # Verificar si tiene imágenes
        for page_num in range(min(3, len(doc))):  # Revisar solo primeras 3 páginas
            page = doc[page_num]
            image_list = page.get_images()
            if image_list:
                info['has_images'] = True
                break
        
        doc.close()
        return info
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error al leer PDF: {str(e)}{Style.RESET_ALL}")
        return {
            'pages': 0,
            'size': get_file_size(pdf_path),
            'title': 'Error al leer',
            'author': 'Desconocido',
            'has_text': False,
            'has_images': False
        }


def create_output_directory(output_path: str) -> bool:
    """
    Crea el directorio de salida si no existe
    
    Args:
        output_path: Ruta del archivo de salida
    
    Returns:
        bool: True si se creó o ya existía el directorio
    """
    try:
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"{Fore.GREEN}📁 Directorio creado: {output_dir}{Style.RESET_ALL}")
        return True
    except Exception as e:
        print(f"{Fore.RED}❌ Error al crear directorio: {str(e)}{Style.RESET_ALL}")
        return False


def get_pdf_files(directory: str) -> List[str]:
    """
    Obtiene todos los archivos PDF de un directorio
    
    Args:
        directory: Ruta al directorio
    
    Returns:
        List[str]: Lista de archivos PDF encontrados
    """
    pdf_files = []
    
    if not os.path.exists(directory):
        print(f"{Fore.RED}❌ Error: El directorio no existe: {directory}{Style.RESET_ALL}")
        return pdf_files
    
    try:
        for file in os.listdir(directory):
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(directory, file))
        
        if not pdf_files:
            print(f"{Fore.YELLOW}⚠️  No se encontraron archivos PDF en: {directory}{Style.RESET_ALL}")
        
        return sorted(pdf_files)
    
    except Exception as e:
        print(f"{Fore.RED}❌ Error al leer directorio: {str(e)}{Style.RESET_ALL}")
        return []


def generate_output_filename(input_path: str, output_path: str = None) -> str:
    """
    Genera el nombre del archivo de salida
    
    Args:
        input_path: Ruta del archivo de entrada
        output_path: Ruta de salida (opcional)
    
    Returns:
        str: Ruta del archivo de salida
    """
    if output_path:
        return output_path
    
    # Si no se especifica salida, crear en el mismo directorio
    input_dir = os.path.dirname(input_path)
    input_name = os.path.splitext(os.path.basename(input_path))[0]
    
    return os.path.join(input_dir, f"{input_name}.docx")


def print_file_info(pdf_path: str, output_path: str, verbose: bool = False) -> None:
    """
    Muestra información del archivo a procesar
    
    Args:
        pdf_path: Ruta del archivo PDF
        output_path: Ruta del archivo de salida
        verbose: Si mostrar información detallada
    """
    if not verbose:
        return
    
    print(f"\n{Fore.CYAN}📋 Información del archivo:{Style.RESET_ALL}")
    print(f"  📄 Entrada: {pdf_path}")
    print(f"  💾 Salida: {output_path}")
    
    info = get_pdf_info(pdf_path)
    print(f"  📊 Páginas: {info['pages']}")
    print(f"  📏 Tamaño: {info['size']}")
    
    if info['title'] != 'Sin título':
        print(f"  📝 Título: {info['title']}")
    
    # Mostrar capacidades detectadas
    capabilities = []
    if info['has_text']:
        capabilities.append("texto")
    if info['has_images']:
        capabilities.append("imágenes")
    
    if capabilities:
        print(f"  🔍 Contiene: {', '.join(capabilities)}")
    else:
        print(f"  {Fore.YELLOW}⚠️  Advertencia: No se detectó texto seleccionable{Style.RESET_ALL}")


def print_success_message(output_path: str, processing_time: float = None) -> None:
    """
    Muestra mensaje de éxito
    
    Args:
        output_path: Ruta del archivo creado
        processing_time: Tiempo de procesamiento en segundos
    """
    print(f"\n{Fore.GREEN}✅ ¡Conversión completada exitosamente!{Style.RESET_ALL}")
    print(f"📄 Archivo creado: {output_path}")
    print(f"📏 Tamaño: {get_file_size(output_path)}")
    
    if processing_time:
        print(f"⏱️  Tiempo: {processing_time:.2f} segundos")


def print_error_message(error: str) -> None:
    """
    Muestra mensaje de error formateado
    
    Args:
        error: Mensaje de error
    """
    print(f"\n{Fore.RED}❌ Error durante la conversión:{Style.RESET_ALL}")
    print(f"   {error}")
    print(f"\n{Fore.YELLOW}💡 Sugerencias:{Style.RESET_ALL}")
    print("   • Verifica que el archivo PDF no esté dañado")
    print("   • Asegúrate de que el PDF tenga texto seleccionable")
    print("   • Revisa los permisos del archivo")


def is_pdf_readable(pdf_path: str) -> bool:
    """
    Verifica si un PDF se puede leer correctamente
    
    Args:
        pdf_path: Ruta al archivo PDF
    
    Returns:
        bool: True si el PDF es legible
    """
    try:
        doc = fitz.open(pdf_path)
        # Intentar leer la primera página
        if len(doc) > 0:
            page = doc[0]
            text = page.get_text()
            doc.close()
            return True
        doc.close()
        return False
    except Exception:
        return False


def sanitize_filename(filename: str) -> str:
    """
    Limpia un nombre de archivo removiendo caracteres problemáticos
    
    Args:
        filename: Nombre del archivo
    
    Returns:
        str: Nombre limpio
    """
    # Caracteres problemáticos en Windows
    invalid_chars = '<>:"/\\|?*'
    
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    
    # Remover espacios al inicio y final
    filename = filename.strip()
    
    # Limitar longitud
    if len(filename) > 200:
        filename = filename[:200]
    
    return filenamev