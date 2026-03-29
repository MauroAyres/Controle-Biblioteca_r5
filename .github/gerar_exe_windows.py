#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script para gerar executável Windows (.exe) standalone
Execute este script em um computador Windows com Python instalado

Uso: python gerar_exe_windows.py
"""

import subprocess
import sys
import os

def main():
    print("=" * 80)
    print("  GERADOR DE EXECUTÁVEL WINDOWS - CONTROLE DE BIBLIOTECA")
    print("=" * 80)
    print()
    
    # Verificar Python
    print("✓ Python encontrado:", sys.version.split()[0])
    print()
    
    # Instalar PyInstaller
    print("📦 Instalando PyInstaller...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller", "-q"])
        print("✓ PyInstaller instalado!")
    except Exception as e:
        print(f"✗ Erro ao instalar PyInstaller: {e}")
        return False
    
    print()
    
    # Instalar python-docx
    print("📦 Instalando python-docx...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx", "-q"])
        print("✓ python-docx instalado!")
    except Exception as e:
        print(f"✗ Erro ao instalar python-docx: {e}")
        return False
    
    print()
    print("🔨 Gerando executável...")
    print("   Isso pode levar alguns minutos, por favor aguarde...")
    print()
    
    # Gerar executável
    try:
        subprocess.check_call([
            sys.executable, "-m", "PyInstaller",
            "--onefile",
            "--windowed",
            "--name", "Controle_Biblioteca",
            "controle_biblioteca.py"
        ])
        print()
        print("=" * 80)
        print("✓ SUCESSO! Executável gerado com sucesso!")
        print("=" * 80)
        print()
        print("📁 Arquivo gerado em: dist\\Controle_Biblioteca.exe")
        print()
        print("🚀 Próximos passos:")
        print("   1. Abra a pasta 'dist'")
        print("   2. Clique duplo em 'Controle_Biblioteca.exe'")
        print("   3. O aplicativo abrirá!")
        print()
        print("💾 Para distribuir:")
        print("   - Copie o arquivo: dist\\Controle_Biblioteca.exe")
        print("   - Envie para outros computadores")
        print("   - Clique duplo para executar (sem precisar de Python!)")
        print()
        print("=" * 80)
        return True
    except Exception as e:
        print(f"✗ Erro ao gerar executável: {e}")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
