#!/usr/bin/env python3
"""
afegir_dia.py - Ajuda per a la carrega manual (Assemblea Bombers)
Afegeix un dia a l'historic a partir d'un Excel del TRI descarregat.
Us: python afegir_dia.py "Personal de guardia.xlsx" 2026-08-02
"""
import sys
import os
import subprocess


def main():
    if len(sys.argv) >= 2:
        xlsx = sys.argv[1]
    else:
        xlsx = input("Arrossega aqui el fitxer Excel del TRI i prem Enter:\n> ").strip().strip('"').strip("'")
    if not os.path.exists(xlsx):
        print("\nNo trobo el fitxer: " + xlsx)
        input("\nPrem Enter per sortir.")
        return
    if len(sys.argv) >= 3:
        data = sys.argv[2]
    else:
        data = input("\nQuina data te aquest TRI? (AAAA-MM-DD, o buit per usar la de l'Excel):\n> ").strip()
    script = os.path.join(os.path.dirname(__file__), "parse_tri.py")
    cmd = [sys.executable, script, "--input", xlsx]
    if data:
        cmd += ["--date", data]
    print("\nProcessant...\n")
    subprocess.run(cmd)
    print("\nFet! Si tot ha anat be, la data ja es a data/history.json.")
    input("\nPrem Enter per sortir.")


if __name__ == "__main__":
    main()
