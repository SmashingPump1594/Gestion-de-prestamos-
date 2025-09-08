#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sistema de Gestión de Préstamos de Equipos de Cómputo
Para Prepa 10 - Versión 1.0

Sistema de Gestión
Fecha: 04/09/2025
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
from datetime import datetime, timedelta
import pandas as pd
from typing import Dict, List, Optional

class SistemaPrestamos:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Gestión de Préstamos - Prepa 10")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f0f0f0')

        # Carpeta base del script
        self.base_dir = os.path.dirname(os.path.abspath(__file__))

        # Configurar estilo
        self.setup_styles()

        # Datos del sistema
        self.prestamos = []
        self.equipos = []
        self.usuarios = []
        self.prestamistas = []

        # Cargar datos existentes
        self.cargar_datos()

        # Crear interfaz
        self.crear_interfaz()
        
    def setup_styles(self):
        """Configurar estilos para la interfaz"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Colores personalizados
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'), foreground='#2c3e50')
        style.configure('Header.TLabel', font=('Arial', 12, 'bold'), foreground='#34495e')
        style.configure('Custom.TButton', font=('Arial', 10, 'bold'))
        
    def cargar_datos(self):
        """Cargar datos desde archivos JSON"""
        try:
            # Helper para ruta absoluta
            def ruta_json(nombre):
                return os.path.join(self.base_dir, nombre)

            # Cargar préstamos
            if os.path.exists(ruta_json('prestamos.json')):
                with open(ruta_json('prestamos.json'), 'r', encoding='utf-8') as f:
                    self.prestamos = json.load(f)
            else:
                self.prestamos = []

            # Cargar equipos
            if os.path.exists(ruta_json('equipos.json')):
                with open(ruta_json('equipos.json'), 'r', encoding='utf-8') as f:
                    self.equipos = json.load(f)
            else:
                self.equipos = []

            # Cargar controles
            if os.path.exists(ruta_json('controles.json')):
                with open(ruta_json('controles.json'), 'r', encoding='utf-8') as f:
                    self.controles = json.load(f)
            else:
                self.controles = []

            # Cargar cables
            if os.path.exists(ruta_json('cables.json')):
                with open(ruta_json('cables.json'), 'r', encoding='utf-8') as f:
                    self.cables = json.load(f)
            else:
                self.cables = []

            # Cargar audifonos
            if os.path.exists(ruta_json('audifonos.json')):
                with open(ruta_json('audifonos.json'), 'r', encoding='utf-8') as f:
                    self.audifonos = json.load(f)
            else:
                self.audifonos = []

            # Cargar usuarios
            if os.path.exists(ruta_json('usuarios.json')):
                with open(ruta_json('usuarios.json'), 'r', encoding='utf-8') as f:
                    self.usuarios = json.load(f)
            else:
                self.usuarios = []

            # Cargar prestamistas
            if os.path.exists(ruta_json('prestamistas.json')):
                with open(ruta_json('prestamistas.json'), 'r', encoding='utf-8') as f:
                    self.prestamistas = json.load(f)
            else:
                self.prestamistas = []

        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar datos: {str(e)}")
    
    def guardar_datos(self):
        """Guardar datos en archivos JSON"""
        try:
            def ruta_json(nombre):
                return os.path.join(self.base_dir, nombre)

            with open(ruta_json('prestamos.json'), 'w', encoding='utf-8') as f:
                json.dump(self.prestamos, f, ensure_ascii=False, indent=2)
            with open(ruta_json('equipos.json'), 'w', encoding='utf-8') as f:
                json.dump(self.equipos, f, ensure_ascii=False, indent=2)
            with open(ruta_json('controles.json'), 'w', encoding='utf-8') as f:
                json.dump(self.controles, f, ensure_ascii=False, indent=2)
            with open(ruta_json('cables.json'), 'w', encoding='utf-8') as f:
                json.dump(self.cables, f, ensure_ascii=False, indent=2)
            with open(ruta_json('audifonos.json'), 'w', encoding='utf-8') as f:
                json.dump(self.audifonos, f, ensure_ascii=False, indent=2)
            with open(ruta_json('usuarios.json'), 'w', encoding='utf-8') as f:
                json.dump(self.usuarios, f, ensure_ascii=False, indent=2)
            with open(ruta_json('prestamistas.json'), 'w', encoding='utf-8') as f:
                json.dump(self.prestamistas, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar datos: {str(e)}")
    
    def crear_interfaz(self):
        """Crear la interfaz principal"""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Título
        title_label = ttk.Label(main_frame, text="Gestión de Préstamos", 
                               style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Crear notebook para pestañas
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Pestañas
        self.crear_pestana_prestamos()
        self.crear_pestana_inventario()
        self.crear_pestana_usuarios()
        self.crear_pestana_reportes()
        
        # Configurar grid para expansión
        main_frame.rowconfigure(1, weight=1)
        
    def crear_pestana_prestamos(self):
        """Crear pestaña de gestión de préstamos"""
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Gestión de Préstamos")
        
        # Frame para formulario de préstamo
        form_frame = ttk.LabelFrame(frame, text="Nuevo Préstamo", padding="10")
        form_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Campos del formulario
        ttk.Label(form_frame, text="Usuario que solicita:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.usuario_var = tk.StringVar()
        self.usuario_combo = ttk.Combobox(form_frame, textvariable=self.usuario_var, width=30)
        self.usuario_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=2)
        self.usuario_combo.bind('<KeyRelease>', self.auto_completar_usuario)
        
        ttk.Label(form_frame, text="Prestamista (quien entrega):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.prestamista_var = tk.StringVar()
        self.prestamista_combo = ttk.Combobox(form_frame, textvariable=self.prestamista_var, width=30)
        self.prestamista_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=2)
        self.prestamista_combo.bind('<KeyRelease>', self.auto_completar_prestamista)
        
        ttk.Label(form_frame, text="Equipo:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.equipo_var = tk.StringVar()
        self.equipo_combo = ttk.Combobox(form_frame, textvariable=self.equipo_var, width=30)
        self.equipo_combo.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=2)

        ttk.Label(form_frame, text="Controles:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.controles_var = tk.StringVar()
        self.controles_combo = ttk.Combobox(form_frame, textvariable=self.controles_var, width=30)
        self.controles_combo.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=2)

        ttk.Label(form_frame, text="Cables:").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.cables_var = tk.StringVar()
        self.cables_combo = ttk.Combobox(form_frame, textvariable=self.cables_var, width=30)
        self.cables_combo.grid(row=4, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=2)

        ttk.Label(form_frame, text="Audifonos:").grid(row=5, column=0, sticky=tk.W, pady=2)
        self.audifonos_var = tk.StringVar()
        self.audifonos_combo = ttk.Combobox(form_frame, textvariable=self.audifonos_var, width=30)
        self.audifonos_combo.grid(row=5, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=2)
        
        ttk.Label(form_frame, text="Estado del equipo:").grid(row=6, column=0, sticky=tk.W, pady=2)
        self.estado_var = tk.StringVar(value="Completo")
        estado_frame = ttk.Frame(form_frame)
        estado_frame.grid(row=6, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=2)
        ttk.Radiobutton(estado_frame, text="Completo", variable=self.estado_var, value="Completo").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(estado_frame, text="Faltante", variable=self.estado_var, value="Faltante").pack(side=tk.LEFT)
        
        ttk.Label(form_frame, text="Observaciones:").grid(row=7, column=0, sticky=tk.W, pady=2)
        self.observaciones_text = tk.Text(form_frame, height=3, width=40)
        self.observaciones_text.grid(row=7, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=2)
        
        # Botones
        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=8, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Registrar Préstamo", 
                  command=self.registrar_prestamo, style='Custom.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Marcar como Entregado", 
                  command=self.marcar_entregado, style='Custom.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Eliminar Préstamo", 
                  command=self.eliminar_prestamo, style='Custom.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Limpiar Formulario", 
                  command=self.limpiar_formulario).pack(side=tk.LEFT, padx=5)
        
        # Frame para lista de préstamos
        list_frame = ttk.LabelFrame(frame, text="Préstamos Actuales", padding="10")
        list_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # Treeview para mostrar préstamos

        columns = ('ID', 'Usuario', 'Prestamista', 'Equipo', 'Controles', 'Cables', 'Audifonos', 'Estado Equipo', 'Observaciones', 'Fecha Préstamo', 'Fecha Entrega', 'Quien Recibe', 'Estado', 'Observaciones Finales')
        self.prestamos_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=15)

        # Configurar columnas
        for col in columns:
            if col == 'Observaciones':
                self.prestamos_tree.heading(col, text=col)
                self.prestamos_tree.column(col, width=180)
            else:
                self.prestamos_tree.heading(col, text=col)
                self.prestamos_tree.column(col, width=120)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.prestamos_tree.yview)
        self.prestamos_tree.configure(yscrollcommand=scrollbar.set)
        
        self.prestamos_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configurar grid
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        
        # Actualizar listas desplegables
        self.actualizar_listas_desplegables()
        self.actualizar_lista_prestamos()
    
    def crear_pestana_inventario(self):
        """Crear pestaña de gestión de inventario"""
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Inventario de Equipos")
        
        # Frame para agregar equipo
        add_frame = ttk.LabelFrame(frame, text="Agregar Nuevo Equipo", padding="10")
        add_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(add_frame, text="Nombre del equipo:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.nuevo_equipo_var = tk.StringVar()
        ttk.Entry(add_frame, textvariable=self.nuevo_equipo_var, width=30).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=2)
        
        ttk.Label(add_frame, text="Categoría:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.categoria_var = tk.StringVar()
        categoria_combo = ttk.Combobox(add_frame, textvariable=self.categoria_var, width=30)
        categoria_combo['values'] = ('Computadora', 'Audifonos', 'Cable', 'Controles')
        categoria_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=2)
        
        button_frame_equipos = ttk.Frame(add_frame)
        button_frame_equipos.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame_equipos, text="Agregar Equipo", 
                  command=self.agregar_equipo, style='Custom.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame_equipos, text="Eliminar Equipo", 
                  command=self.eliminar_equipo, style='Custom.TButton').pack(side=tk.LEFT, padx=5)
        
        # Frame para lista de equipos
        list_frame = ttk.LabelFrame(frame, text="Equipos Disponibles", padding="10")
        list_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # Treeview para equipos
        columns = ('ID', 'Nombre', 'Categoría', 'Estado')
        self.equipos_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=15)
        
        for col in columns:
            self.equipos_tree.heading(col, text=col)
            self.equipos_tree.column(col, width=150)
        
        scrollbar_equipos = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.equipos_tree.yview)
        self.equipos_tree.configure(yscrollcommand=scrollbar_equipos.set)
        
        self.equipos_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_equipos.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configurar grid
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        
        self.actualizar_lista_equipos()
    
    def crear_pestana_usuarios(self):
        """Crear pestaña de gestión de usuarios"""
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Gestión de Usuarios")
        
        # Frame para usuarios
        usuarios_frame = ttk.LabelFrame(frame, text="Usuarios (Solicitantes)", padding="10")
        usuarios_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        
        ttk.Label(usuarios_frame, text="Agregar Usuario:").pack(anchor=tk.W)
        self.nuevo_usuario_var = tk.StringVar()
        ttk.Entry(usuarios_frame, textvariable=self.nuevo_usuario_var, width=30).pack(fill=tk.X, pady=(0, 5))
        ttk.Button(usuarios_frame, text="Agregar", 
                  command=self.agregar_usuario, style='Custom.TButton').pack(pady=(0, 10))
        
        # Lista de usuarios
        self.usuarios_listbox = tk.Listbox(usuarios_frame, height=15)
        self.usuarios_listbox.pack(fill=tk.BOTH, expand=True)
        ttk.Button(usuarios_frame, text="Eliminar Seleccionado", 
                  command=self.eliminar_usuario).pack(pady=(5, 0))
        
        # Frame para prestamistas
        prestamistas_frame = ttk.LabelFrame(frame, text="Prestamistas (Entregan Equipos)", padding="10")
        prestamistas_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        
        ttk.Label(prestamistas_frame, text="Agregar Prestamista:").pack(anchor=tk.W)
        self.nuevo_prestamista_var = tk.StringVar()
        ttk.Entry(prestamistas_frame, textvariable=self.nuevo_prestamista_var, width=30).pack(fill=tk.X, pady=(0, 5))
        ttk.Button(prestamistas_frame, text="Agregar", 
                  command=self.agregar_prestamista, style='Custom.TButton').pack(pady=(0, 10))
        
        # Lista de prestamistas
        self.prestamistas_listbox = tk.Listbox(prestamistas_frame, height=15)
        self.prestamistas_listbox.pack(fill=tk.BOTH, expand=True)
        ttk.Button(prestamistas_frame, text="Eliminar Seleccionado", 
                  command=self.eliminar_prestamista).pack(pady=(5, 0))
        
        # Configurar grid
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(0, weight=1)
        
        self.actualizar_listas_usuarios()
    
    def crear_pestana_reportes(self):
        """Crear pestaña de reportes y exportación"""
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Reportes y Exportación")
        
        # Frame para búsquedas
        search_frame = ttk.LabelFrame(frame, text="Búsquedas", padding="10")
        search_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(search_frame, text="Buscar por:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.busqueda_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.busqueda_var, width=30).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=2)
        
        ttk.Label(search_frame, text="Tipo:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.tipo_busqueda_var = tk.StringVar(value="Usuario")
        tipo_combo = ttk.Combobox(search_frame, textvariable=self.tipo_busqueda_var, width=30)
        tipo_combo['values'] = ('Usuario', 'Prestamista', 'Equipo', 'Estado')
        tipo_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=2)
        
        ttk.Button(search_frame, text="Buscar", 
                  command=self.buscar_prestamos, style='Custom.TButton').grid(row=2, column=0, columnspan=2, pady=10)
        
        # Frame para resultados
        results_frame = ttk.LabelFrame(frame, text="Resultados de Búsqueda", padding="10")
        results_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # Treeview para resultados
        columns = ('ID', 'Usuario', 'Prestamista', 'Equipo', 'Estado Equipo', 'Observaciones', 'Fecha Préstamo', 'Quien Recibe', 'Estado','Fecha Entrega', 'Observaciones_Finales')
        self.resultados_tree = ttk.Treeview(results_frame, columns=columns, show='headings', height=10)
        
        for col in columns:
            self.resultados_tree.heading(col, text=col)
            self.resultados_tree.column(col, width=120)
        
        scrollbar_results = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.resultados_tree.yview)
        self.resultados_tree.configure(yscrollcommand=scrollbar_results.set)
        
        self.resultados_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_results.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Frame para exportación
        export_frame = ttk.LabelFrame(frame, text="Exportar Datos", padding="10")
        export_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Button(export_frame, text="Exportar a Excel", 
                  command=self.exportar_excel, style='Custom.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(export_frame, text="Exportar Préstamos Activos", 
                  command=self.exportar_prestamos_activos, style='Custom.TButton').pack(side=tk.LEFT, padx=5)
        
        # Configurar grid
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
    
    def actualizar_listas_desplegables(self):
        """Listas desplegables"""
        # Actualizar usuarios
        usuarios_nombres = [f"{u['nombre']}" for u in self.usuarios]
        self.usuario_combo['values'] = usuarios_nombres

        # Actualizar prestamistas
        prestamistas_nombres = [f"{p['nombre']}" for p in self.prestamistas]
        self.prestamista_combo['values'] = prestamistas_nombres

        # Actualizar equipos disponibles (incluye controles, cables y audífonos)
        equipos_disponibles = [f"{e['nombre']} ({e['categoria']})" for e in self.equipos if e['estado'] == 'Disponible']
        controles_disponibles = [f"{c['nombre']} (Controles)" for c in getattr(self, 'controles', []) if c.get('estado', 'Disponible') == 'Disponible']
        cables_disponibles = [f"{c['nombre']} (Cable)" for c in getattr(self, 'cables', []) if c.get('estado', 'Disponible') == 'Disponible']
        audifonos_disponibles = [f"{a['nombre']} (Audifonos)" for a in getattr(self, 'audifonos', []) if a.get('estado', 'Disponible') == 'Disponible']
        self.equipo_combo['values'] = equipos_disponibles + controles_disponibles + cables_disponibles + audifonos_disponibles

        # Actualizar controles disponibles
        self.controles_combo['values'] = [f"{c['nombre']}" for c in getattr(self, 'controles', []) if c.get('estado', 'Disponible') == 'Disponible']

        # Actualizar cables disponibles
        self.cables_combo['values'] = [f"{c['nombre']}" for c in getattr(self, 'cables', []) if c.get('estado', 'Disponible') == 'Disponible']

        # Actualizar audífonos disponibles
        self.audifonos_combo['values'] = [f"{a['nombre']}" for a in getattr(self, 'audifonos', []) if a.get('estado', 'Disponible') == 'Disponible']
    
    def auto_completar_usuario(self, event):
        """Auto-completar usuario mientras escribe"""
        texto = self.usuario_var.get().lower()
        if texto:
            coincidencias = [u for u in self.usuarios if texto in u['nombre'].lower()]
            self.usuario_combo['values'] = [u['nombre'] for u in coincidencias]
        else:
            self.usuario_combo['values'] = [u['nombre'] for u in self.usuarios]
    
    def auto_completar_prestamista(self, event):
        """Auto-completar prestamista mientras escribe"""
        texto = self.prestamista_var.get().lower()
        if texto:
            coincidencias = [p for p in self.prestamistas if texto in p['nombre'].lower()]
            self.prestamista_combo['values'] = [p['nombre'] for p in coincidencias]
        else:
            self.prestamista_combo['values'] = [p['nombre'] for p in self.prestamistas]
    
    def actualizar_lista_prestamos(self):
        """Actualizar la lista de préstamos en el treeview"""
        # Limpiar treeview
        for item in self.prestamos_tree.get_children():
            self.prestamos_tree.delete(item)
        
        # Agregar préstamos
        for prestamo in self.prestamos:
            self.prestamos_tree.insert('', 'end', values=(
                prestamo['id'],
                prestamo['usuario'],
                prestamo['prestamista'],
                prestamo['equipo'],
                prestamo.get('controles', ''),
                prestamo.get('cables', ''),
                prestamo.get('audifonos', ''),
                prestamo.get('estado_equipo', 'Completo'),
                prestamo.get('observaciones', ''),
                prestamo['fecha_prestamo'],
                prestamo.get('fecha_entrega', 'Pendiente'),
                prestamo.get('quien_recibe', ''),
                prestamo['estado'],
                prestamo.get('observaciones_finales', '')
            ))
    
    def actualizar_lista_equipos(self):
        """Actualizar la lista de equipos en el treeview"""
        # Limpiar treeview
        for item in self.equipos_tree.get_children():
            self.equipos_tree.delete(item)
        
        # Agregar equipos
        for equipo in self.equipos:
            self.equipos_tree.insert('', 'end', values=(
                equipo.get('id', ''),
                equipo.get('nombre', ''),
                equipo.get('categoria', ''),
                equipo.get('estado', '')
            ))
        # Agregar controles
        for control in getattr(self, 'controles', []):
            self.equipos_tree.insert('', 'end', values=(
                control.get('id', ''),
                control.get('nombre', ''),
                'Controles',
                control.get('estado', '')
            ))
        # Agregar cables
        for cable in getattr(self, 'cables', []):
            self.equipos_tree.insert('', 'end', values=(
                cable.get('id', ''),
                cable.get('nombre', ''),
                'Cable',
                cable.get('estado', '')
            ))
        # Agregar audífonos
        for audifono in getattr(self, 'audifonos', []):
            self.equipos_tree.insert('', 'end', values=(
                audifono.get('id', ''),
                audifono.get('nombre', ''),
                'Audifonos',
                audifono.get('estado', '')
            ))

    
    def actualizar_listas_usuarios(self):
        """Actualizar las listas de usuarios y prestamistas"""
        # Limpiar listboxes
        self.usuarios_listbox.delete(0, tk.END)
        self.prestamistas_listbox.delete(0, tk.END)
        
        # Agregar usuarios
        for usuario in self.usuarios:
            self.usuarios_listbox.insert(tk.END, usuario['nombre'])
        
        # Agregar prestamistas
        for prestamista in self.prestamistas:
            self.prestamistas_listbox.insert(tk.END, prestamista['nombre'])
    
    def registrar_prestamo(self):
        """Registrar un nuevo préstamo"""
        try:
            # Validar campos
            if not self.usuario_var.get():
                messagebox.showwarning("Advertencia", "Se recomienda llenar al menos Usuario y Equipo")

            # Verificar si el usuario existe, si no, agregarlo
            usuario_nombre = self.usuario_var.get().strip()
            usuario_existe = any(u['nombre'] == usuario_nombre for u in self.usuarios)
            if not usuario_existe:
                nuevo_usuario = {
                    'id': max([u['id'] for u in self.usuarios], default=0) + 1,
                    'nombre': usuario_nombre,
                    'tipo': 'Usuario'
                }
                self.usuarios.append(nuevo_usuario)
                messagebox.showinfo("Información", f"Usuario '{usuario_nombre}' agregado automáticamente")

            # Verificar si el prestamista existe, si no, agregarlo
            prestamista_nombre = self.prestamista_var.get().strip()
            prestamista_existe = any(p['nombre'] == prestamista_nombre for p in self.prestamistas)
            if not prestamista_existe:
                nuevo_prestamista = {
                    'id': max([p['id'] for p in self.prestamistas], default=0) + 1,
                    'nombre': prestamista_nombre,
                    'tipo': 'Prestamista'
                }
                self.prestamistas.append(nuevo_prestamista)
                messagebox.showinfo("Información", f"Prestamista '{prestamista_nombre}' agregado automáticamente")

            # Obtener equipos seleccionados
            equipo_seleccionado = self.equipo_var.get()
            control_seleccionado = self.controles_var.get()
            cable_seleccionado = self.cables_var.get()
            audifono_seleccionado = self.audifonos_var.get()
            
            # Lista para almacenar todos los equipos seleccionados
            equipos_seleccionados = []
            
            # Función auxiliar para buscar un equipo en las listas
            def buscar_equipo(nombre_buscado, categoria_buscada):
                if categoria_buscada == 'Computadora':
                    for equipo in self.equipos:
                        if f"{equipo['nombre']} ({equipo['categoria']})" == nombre_buscado:
                            return equipo
                elif categoria_buscada == 'Controles':
                    for control in getattr(self, 'controles', []):
                        if f"{control['nombre']} (Controles)" == nombre_buscado:
                            return control
                elif categoria_buscada == 'Cable':
                    for cable in getattr(self, 'cables', []):
                        if f"{cable['nombre']} (Cable)" == nombre_buscado:
                            return cable
                elif categoria_buscada == 'Audifonos':
                    for audifono in getattr(self, 'audifonos', []):
                        if f"{audifono['nombre']} (Audifonos)" == nombre_buscado:
                            return audifono
                return None
            
            # Buscar equipo principal (computadora)
            if equipo_seleccionado:
                equipo_info = buscar_equipo(equipo_seleccionado, 'Computadora')
                if equipo_info and equipo_info['estado'] == 'Disponible':
                    equipos_seleccionados.append(('equipo', equipo_info))
                elif equipo_info and equipo_info['estado'] != 'Disponible':
                    messagebox.showerror("Error", f"El equipo '{equipo_info['nombre']}' no está disponible")
                    return
                elif not equipo_info:
                    messagebox.showerror("Error", "Equipo no encontrado")
                    return
            
            # Buscar control
            if control_seleccionado:
                control_info = buscar_equipo(f"{control_seleccionado} (Controles)", 'Controles')
                if control_info and control_info['estado'] == 'Disponible':
                    equipos_seleccionados.append(('controles', control_info))
                elif control_info and control_info['estado'] != 'Disponible':
                    messagebox.showerror("Error", f"El control '{control_info['nombre']}' no está disponible")
                    return
                elif not control_info:
                    messagebox.showerror("Error", "Control no encontrado")
                    return
            
            # Buscar cable
            if cable_seleccionado:
                cable_info = buscar_equipo(f"{cable_seleccionado} (Cable)", 'Cable')
                if cable_info and cable_info['estado'] == 'Disponible':
                    equipos_seleccionados.append(('cables', cable_info))
                elif cable_info and cable_info['estado'] != 'Disponible':
                    messagebox.showerror("Error", f"El cable '{cable_info['nombre']}' no está disponible")
                    return
                elif not cable_info:
                    messagebox.showerror("Error", "Cable no encontrado")
                    return
            
            # Buscar audífonos
            if audifono_seleccionado:
                audifono_info = buscar_equipo(f"{audifono_seleccionado} (Audifonos)", 'Audifonos')
                if audifono_info and audifono_info['estado'] == 'Disponible':
                    equipos_seleccionados.append(('audifonos', audifono_info))
                elif audifono_info and audifono_info['estado'] != 'Disponible':
                    messagebox.showerror("Error", f"Los audífonos '{audifono_info['nombre']}' no están disponibles")
                    return
                elif not audifono_info:
                    messagebox.showerror("Error", "Audífonos no encontrados")
                    return
            
            # Verificar que al menos se haya seleccionado un equipo
            if not equipos_seleccionados:
                messagebox.showerror("Error", "Debe seleccionar al menos un equipo")
                return

            # Crear nuevo préstamo con todos los equipos seleccionados
            nuevo_prestamo = {
                'id': max([p['id'] for p in self.prestamos], default=0) + 1,
                'usuario': usuario_nombre,
                'prestamista': prestamista_nombre,
                'equipo': '',  # Se llenará con el equipo principal si existe
                'controles': '',  # Se llenará con el control si existe
                'cables': '',  # Se llenará con el cable si existe
                'audifonos': '',  # Se llenará con los audífonos si existen
                'estado_equipo': self.estado_var.get(),
                'fecha_prestamo': datetime.now().strftime("%Y-%m-%d %H:%M"),
                'fecha_entrega': None,
                'quien_recibe': '',
                'estado': 'Prestado',
                'observaciones': self.observaciones_text.get("1.0", tk.END).strip()
            }
            
            # Llenar los campos de equipos según lo seleccionado
            for tipo_equipo, equipo_info in equipos_seleccionados:
                if tipo_equipo == 'equipo':
                    nuevo_prestamo['equipo'] = equipo_info['nombre']
                elif tipo_equipo == 'controles':
                    nuevo_prestamo['controles'] = equipo_info['nombre']
                elif tipo_equipo == 'cables':
                    nuevo_prestamo['cables'] = equipo_info['nombre']
                elif tipo_equipo == 'audifonos':
                    nuevo_prestamo['audifonos'] = equipo_info['nombre']

            # Agregar préstamo
            self.prestamos.append(nuevo_prestamo)

            # Actualizar estado de todos los equipos seleccionados
            for tipo_equipo, equipo_info in equipos_seleccionados:
                equipo_info['estado'] = 'Prestado'

            # Guardar datos
            self.guardar_datos()

            # Actualizar interfaces
            self.actualizar_lista_prestamos()
            self.actualizar_lista_equipos()
            self.actualizar_listas_desplegables()

            # Limpiar formulario
            self.limpiar_formulario()

            messagebox.showinfo("Éxito", "Préstamo registrado correctamente")

        except Exception as e:
            messagebox.showerror("Error", f"Error al registrar préstamo: {str(e)}")
    
    def marcar_entregado(self):
        """Marcar un equipo como entregado"""
        try:
            # Obtener selección del treeview
            seleccion = self.prestamos_tree.selection()
            if not seleccion:
                messagebox.showwarning("Advertencia", "Por favor seleccione un préstamo para marcar como entregado")
                return
            
            # Obtener datos del préstamo
            item = self.prestamos_tree.item(seleccion[0])
            prestamo_id = int(item['values'][0])
            
            # Buscar préstamo
            prestamo = None
            for p in self.prestamos:
                if p['id'] == prestamo_id:
                    prestamo = p
                    break
            
            if not prestamo:
                messagebox.showerror("Error", "Préstamo no encontrado")
                return
            
            if prestamo['estado'] == 'Entregado':
                messagebox.showwarning("Advertencia", "Este equipo ya fue marcado como entregado")
                return
            
            # Ventana para capturar quien recibe
            ventana_entrega = tk.Toplevel(self.root)
            ventana_entrega.title("Marcar como Entregado")
            ventana_entrega.geometry("450x320")
            ventana_entrega.transient(self.root)
            ventana_entrega.grab_set()
            
            # Centrar ventana
            ventana_entrega.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
            
            # Frame principal para mejor organización
            main_frame_entrega = ttk.Frame(ventana_entrega, padding="20")
            main_frame_entrega.pack(fill=tk.BOTH, expand=True)
            
            ttk.Label(main_frame_entrega, text=f"Equipo: {prestamo['equipo']}", font=('Arial', 12, 'bold')).pack(pady=(0, 15))
            ttk.Label(main_frame_entrega, text="¿Quién recibe el equipo?").pack(pady=(0, 5))
            
            quien_recibe_var = tk.StringVar()
            quien_recibe_combo = ttk.Combobox(main_frame_entrega, textvariable=quien_recibe_var, width=35)
            quien_recibe_combo['values'] = [p['nombre'] for p in self.prestamistas]
            quien_recibe_combo.pack(pady=(0, 15))
            
            ttk.Label(main_frame_entrega, text="Estado del equipo al entregar:").pack(pady=(0, 5))
            estado_entrega_var = tk.StringVar(value="Completo")
            estado_frame = ttk.Frame(main_frame_entrega)
            estado_frame.pack(pady=(0, 20))
            ttk.Radiobutton(estado_frame, text="Completo", variable=estado_entrega_var, value="Completo").pack(side=tk.LEFT, padx=5)
            ttk.Radiobutton(estado_frame, text="Incompleto", variable=estado_entrega_var, value="Incompleto").pack(side=tk.LEFT, padx=5)

            # Campo de observaciones, oculto por defecto, pero el espacio está reservado
            observaciones_frame = ttk.Frame(main_frame_entrega)
            observaciones_label = ttk.Label(observaciones_frame, text="Observaciones (si falta algo):")
            observaciones_text = tk.Text(observaciones_frame, height=3, width=40)
            observaciones_label.pack(anchor=tk.W, pady=(0, 2))
            observaciones_text.pack(fill=tk.X)
            observaciones_frame.pack(fill=tk.X, pady=(0, 10))
            # Mostrar/ocultar solo el contenido, no el frame
            def mostrar_observaciones(*args):
                if estado_entrega_var.get() == "Incompleto":
                    observaciones_label.configure(state='normal')
                    observaciones_text.configure(state='normal')
                else:
                    observaciones_label.configure(state='disabled')
                    observaciones_text.delete("1.0", tk.END)
                    observaciones_text.configure(state='disabled')
            estado_entrega_var.trace_add('write', mostrar_observaciones)
            mostrar_observaciones()
            
            def confirmar_entrega():
                if not quien_recibe_var.get().strip():
                    messagebox.showerror("Error", "Por favor especifique quién recibe el equipo")
                    return

                # Observaciones si incompleto
                observaciones = ""
                if estado_entrega_var.get() == "Incompleto":
                    observaciones = observaciones_text.get("1.0", tk.END).strip()
                    # Guardar en observaciones.json
                    try:
                        obs_path = os.path.join(self.base_dir, 'observaciones_finales   .json')
                        if os.path.exists(obs_path):
                            with open(obs_path, 'r', encoding='utf-8') as f:
                                obs_data = json.load(f)
                        else:
                            obs_data = []
                        obs_data.append({
                            'prestamo_id': prestamo['id'],
                            'fecha': datetime.now().strftime("%Y-%m-%d %H:%M"),
                            'observaciones': observaciones
                        })
                        with open(obs_path, 'w', encoding='utf-8') as f:
                            json.dump(obs_data, f, ensure_ascii=False, indent=2)
                    except Exception as e:
                        messagebox.showerror("Error", f"No se pudo guardar observaciones: {str(e)}")

                # Actualizar préstamo
                prestamo['fecha_entrega'] = datetime.now().strftime("%Y-%m-%d %H:%M")
                prestamo['quien_recibe'] = quien_recibe_var.get().strip()
                prestamo['estado'] = 'Entregado'
                prestamo['estado_equipo_entrega'] = estado_entrega_var.get()
                if observaciones:
                        prestamo['observaciones_finales'] = observaciones

                # Actualizar estado de todos los equipos del préstamo
                # Equipo principal
                if prestamo.get('equipo'):
                    for equipo in self.equipos:
                        if equipo['nombre'] == prestamo['equipo']:
                            equipo['estado'] = 'Disponible'
                            break

                # Controles
                if prestamo.get('controles'):
                    for control in getattr(self, 'controles', []):
                        if control['nombre'] == prestamo['controles']:
                            control['estado'] = 'Disponible'
                            break

                # Cables
                if prestamo.get('cables'):
                    for cable in getattr(self, 'cables', []):
                        if cable['nombre'] == prestamo['cables']:
                            cable['estado'] = 'Disponible'
                            break

                # Audífonos
                if prestamo.get('audifonos'):
                    for audifono in getattr(self, 'audifonos', []):
                        if audifono['nombre'] == prestamo['audifonos']:
                            audifono['estado'] = 'Disponible'
                            break

                # Guardar datos
                self.guardar_datos()

                # Actualizar interfaces
                self.actualizar_lista_prestamos()
                self.actualizar_lista_equipos()
                self.actualizar_listas_desplegables()

                ventana_entrega.destroy()
                messagebox.showinfo("Éxito", "Equipo marcado como entregado correctamente")
            
            # Frame para botones
            button_frame_entrega = ttk.Frame(main_frame_entrega)
            button_frame_entrega.pack(fill=tk.X)
            
            ttk.Button(button_frame_entrega, text="Confirmar Entrega", 
                      command=confirmar_entrega, style='Custom.TButton').pack(side=tk.LEFT, padx=(0, 10))
            ttk.Button(button_frame_entrega, text="Cancelar", 
                      command=ventana_entrega.destroy).pack(side=tk.LEFT)   
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al marcar como entregado: {str(e)}")
    
    def eliminar_prestamo(self):
        """Eliminar préstamo seleccionado"""
        try:
            # Obtener selección del treeview
            seleccion = self.prestamos_tree.selection()
            if not seleccion:
                messagebox.showwarning("Advertencia", "Por favor seleccione un préstamo para eliminar")
                return
            
            # Obtener datos del préstamo
            item = self.prestamos_tree.item(seleccion[0])
            prestamo_id = int(item['values'][0])
            
            # Buscar préstamo
            prestamo = None
            for p in self.prestamos:
                if p['id'] == prestamo_id:
                    prestamo = p
                    break
            
            if not prestamo:
                messagebox.showerror("Error", "Préstamo no encontrado")
                return
            
            # Crear descripción del préstamo para el mensaje
            equipos_desc = []
            if prestamo.get('equipo'):
                equipos_desc.append(f"Equipo: {prestamo['equipo']}")
            if prestamo.get('controles'):
                equipos_desc.append(f"Control: {prestamo['controles']}")
            if prestamo.get('cables'):
                equipos_desc.append(f"Cable: {prestamo['cables']}")
            if prestamo.get('audifonos'):
                equipos_desc.append(f"Audífonos: {prestamo['audifonos']}")
            
            descripcion = ", ".join(equipos_desc) if equipos_desc else "Sin equipos"
            
            # Confirmar eliminación
            if messagebox.askyesno("Confirmar", f"¿Eliminar el préstamo de: {descripcion}?"):
                # Si el préstamo está activo, liberar todos los equipos del préstamo
                if prestamo['estado'] == 'Prestado':
                    # Equipo principal
                    if prestamo.get('equipo'):
                        for equipo in self.equipos:
                            if equipo['nombre'] == prestamo['equipo']:
                                equipo['estado'] = 'Disponible'
                                break
                    
                    # Controles
                    if prestamo.get('controles'):
                        for control in getattr(self, 'controles', []):
                            if control['nombre'] == prestamo['controles']:
                                control['estado'] = 'Disponible'
                                break
                    
                    # Cables
                    if prestamo.get('cables'):
                        for cable in getattr(self, 'cables', []):
                            if cable['nombre'] == prestamo['cables']:
                                cable['estado'] = 'Disponible'
                                break
                    
                    # Audífonos
                    if prestamo.get('audifonos'):
                        for audifono in getattr(self, 'audifonos', []):
                            if audifono['nombre'] == prestamo['audifonos']:
                                audifono['estado'] = 'Disponible'
                                break
                
                # Eliminar préstamo
                self.prestamos.remove(prestamo)
                
                # Guardar datos
                self.guardar_datos()
                
                # Actualizar interfaces
                self.actualizar_lista_prestamos()
                self.actualizar_lista_equipos()
                self.actualizar_listas_desplegables()
                
                messagebox.showinfo("Éxito", "Préstamo eliminado correctamente")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al eliminar préstamo: {str(e)}")
    
    def limpiar_formulario(self):
        """Limpiar el formulario de préstamo"""
        self.usuario_var.set("")
        self.prestamista_var.set("")
        self.equipo_var.set("")
        self.controles_var.set("")
        self.cables_var.set("")
        self.audifonos_var.set("")
        self.estado_var.set("Completo")
        self.observaciones_text.delete("1.0", tk.END)
    
    def agregar_equipo(self):
        """Agregar un nuevo equipo al inventario"""
        try:
            nombre = self.nuevo_equipo_var.get().strip()
            categoria = self.categoria_var.get().strip()
            
            if not nombre or not categoria:
                messagebox.showerror("Error", "Por favor complete todos los campos")
                return
            
            # Crear nuevo equipo
            nuevo_equipo = {
                'id': 0,  # Se calculará según la lista correspondiente
                'nombre': nombre,
                'categoria': categoria,
                'estado': 'Disponible'
            }
            
            # Agregar equipo a la lista correspondiente según la categoría
            if categoria == 'Computadora':
                nuevo_equipo['id'] = max([e['id'] for e in self.equipos], default=0) + 1
                self.equipos.append(nuevo_equipo)
            elif categoria == 'Controles':
                nuevo_equipo['id'] = max([c['id'] for c in getattr(self, 'controles', [])], default=0) + 1
                if not hasattr(self, 'controles'):
                    self.controles = []
                self.controles.append(nuevo_equipo)
            elif categoria == 'Cable':
                nuevo_equipo['id'] = max([c['id'] for c in getattr(self, 'cables', [])], default=0) + 1
                if not hasattr(self, 'cables'):
                    self.cables = []
                self.cables.append(nuevo_equipo)
            elif categoria == 'Audifonos':
                nuevo_equipo['id'] = max([a['id'] for a in getattr(self, 'audifonos', [])], default=0) + 1
                if not hasattr(self, 'audifonos'):
                    self.audifonos = []
                self.audifonos.append(nuevo_equipo)
            else:
                # Por defecto, agregar a equipos
                nuevo_equipo['id'] = max([e['id'] for e in self.equipos], default=0) + 1
                self.equipos.append(nuevo_equipo)
            
            # Guardar datos
            self.guardar_datos()
            
            # Actualizar interfaces
            self.actualizar_lista_equipos()
            self.actualizar_listas_desplegables()
            
            # Limpiar formulario
            self.nuevo_equipo_var.set("")
            self.categoria_var.set("")
            
            messagebox.showinfo("Éxito", "Equipo agregado correctamente")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al agregar equipo: {str(e)}")
    
    def eliminar_equipo(self):
        """Eliminar equipo seleccionado"""
        try:
            # Obtener selección del treeview
            seleccion = self.equipos_tree.selection()
            if not seleccion:
                messagebox.showwarning("Advertencia", "Por favor seleccione un equipo para eliminar")
                return
            
            # Obtener datos del equipo
            item = self.equipos_tree.item(seleccion[0])
            equipo_id = int(item['values'][0])
            
            # Buscar equipo en todas las listas
            equipo = None
            lista_origen = None
            
            # Buscar en equipos
            for e in self.equipos:
                if e['id'] == equipo_id:
                    equipo = e
                    lista_origen = self.equipos
                    break
            
            # Buscar en controles si no se encontró
            if not equipo:
                for c in getattr(self, 'controles', []):
                    if c['id'] == equipo_id:
                        equipo = c
                        lista_origen = getattr(self, 'controles', [])
                        break
            
            # Buscar en cables si no se encontró
            if not equipo:
                for c in getattr(self, 'cables', []):
                    if c['id'] == equipo_id:
                        equipo = c
                        lista_origen = getattr(self, 'cables', [])
                        break
            
            # Buscar en audífonos si no se encontró
            if not equipo:
                for a in getattr(self, 'audifonos', []):
                    if a['id'] == equipo_id:
                        equipo = a
                        lista_origen = getattr(self, 'audifonos', [])
                        break
            
            if not equipo:
                messagebox.showerror("Error", "Equipo no encontrado")
                return
            
            # Verificar si tiene préstamos activos
            prestamos_activos = [p for p in self.prestamos if p['equipo'] == equipo['nombre'] and p['estado'] == 'Prestado']
            if prestamos_activos:
                messagebox.showerror("Error", "No se puede eliminar el equipo porque tiene préstamos activos")
                return
            
            # Confirmar eliminación
            if messagebox.askyesno("Confirmar", f"¿Eliminar el equipo '{equipo['nombre']}'?"):
                # Eliminar equipo de la lista correspondiente
                lista_origen.remove(equipo)
                
                # Guardar datos
                self.guardar_datos()
                
                # Actualizar interfaces
                self.actualizar_lista_equipos()
                self.actualizar_listas_desplegables()
                
                messagebox.showinfo("Éxito", "Equipo eliminado correctamente")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al eliminar equipo: {str(e)}")
    
    def agregar_usuario(self):
        """Agregar un nuevo usuario"""
        try:
            nombre = self.nuevo_usuario_var.get().strip()
            if not nombre:
                messagebox.showerror("Error", "Por favor ingrese el nombre del usuario")
                return
            
            # Crear nuevo usuario
            nuevo_usuario = {
                'id': max([u['id'] for u in self.usuarios], default=0) + 1,
                'nombre': nombre,
                'tipo': 'Usuario'
            }
            
            # Agregar usuario
            self.usuarios.append(nuevo_usuario)
            
            # Guardar datos
            self.guardar_datos()
            
            # Actualizar interfaces
            self.actualizar_listas_usuarios()
            self.actualizar_listas_desplegables()
            
            # Limpiar formulario
            self.nuevo_usuario_var.set("")
            
            messagebox.showinfo("Éxito", "Usuario agregado correctamente")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al agregar usuario: {str(e)}")
    
    def agregar_prestamista(self):
        """Agregar un nuevo prestamista"""
        try:
            nombre = self.nuevo_prestamista_var.get().strip()
            if not nombre:
                messagebox.showerror("Error", "Por favor ingrese el nombre del prestamista")
                return
            
            # Crear nuevo prestamista
            nuevo_prestamista = {
                'id': max([p['id'] for p in self.prestamistas], default=0) + 1,
                'nombre': nombre,
                'tipo': 'Prestamista'
            }
            
            # Agregar prestamista
            self.prestamistas.append(nuevo_prestamista)
            
            # Guardar datos
            self.guardar_datos()
            
            # Actualizar interfaces
            self.actualizar_listas_usuarios()
            self.actualizar_listas_desplegables()
            
            # Limpiar formulario
            self.nuevo_prestamista_var.set("")
            
            messagebox.showinfo("Éxito", "Prestamista agregado correctamente")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al agregar prestamista: {str(e)}")
    
    def eliminar_usuario(self):
        """Eliminar usuario seleccionado"""
        try:
            seleccion = self.usuarios_listbox.curselection()
            if not seleccion:
                messagebox.showwarning("Advertencia", "Por favor seleccione un usuario para eliminar")
                return
            
            indice = seleccion[0]
            usuario = self.usuarios[indice]
            
            if messagebox.askyesno("Confirmar", f"¿Eliminar al usuario '{usuario['nombre']}'?"):
                # Verificar si tiene préstamos activos
                prestamos_activos = [p for p in self.prestamos if p['usuario'] == usuario['nombre'] and p['estado'] == 'Prestado']
                if prestamos_activos:
                    messagebox.showerror("Error", "No se puede eliminar el usuario porque tiene préstamos activos")
                    return
                
                # Eliminar usuario
                del self.usuarios[indice]
                
                # Guardar datos
                self.guardar_datos()
                
                # Actualizar interfaces
                self.actualizar_listas_usuarios()
                self.actualizar_listas_desplegables()
                
                messagebox.showinfo("Éxito", "Usuario eliminado correctamente")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al eliminar usuario: {str(e)}")
    
    def eliminar_prestamista(self):
        """Eliminar prestamista seleccionado"""
        try:
            seleccion = self.prestamistas_listbox.curselection()
            if not seleccion:
                messagebox.showwarning("Advertencia", "Por favor seleccione un prestamista para eliminar")
                return
            
            indice = seleccion[0]
            prestamista = self.prestamistas[indice]
            
            if messagebox.askyesno("Confirmar", f"¿Eliminar al prestamista '{prestamista['nombre']}'?"):
                # Verificar si tiene préstamos activos
                prestamos_activos = [p for p in self.prestamos if p['prestamista'] == prestamista['nombre'] and p['estado'] == 'Prestado']
                if prestamos_activos:
                    messagebox.showerror("Error", "No se puede eliminar el prestamista porque tiene préstamos activos")
                    return
                
                # Eliminar prestamista
                del self.prestamistas[indice]
                
                # Guardar datos
                self.guardar_datos()
                
                # Actualizar interfaces
                self.actualizar_listas_usuarios()
                self.actualizar_listas_desplegables()
                
                messagebox.showinfo("Éxito", "Prestamista eliminado correctamente")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al eliminar prestamista: {str(e)}")
    
    def buscar_prestamos(self):
        """Buscar préstamos según criterios"""
        try:
            busqueda = self.busqueda_var.get().strip().lower()
            tipo = self.tipo_busqueda_var.get()
            
            if not busqueda:
                # Mostrar todos los préstamos
                resultados = self.prestamos
            else:
                # Filtrar según tipo
                resultados = []
                for prestamo in self.prestamos:
                    if tipo == "Usuario" and busqueda in prestamo['usuario'].lower():
                        resultados.append(prestamo)
                    elif tipo == "Prestamista" and busqueda in prestamo['prestamista'].lower():
                        resultados.append(prestamo)
                    elif tipo == "Equipo" and busqueda in prestamo['equipo'].lower():
                        resultados.append(prestamo)
                    elif tipo == "Estado" and busqueda in prestamo['estado'].lower():
                        resultados.append(prestamo)
            
            # Limpiar treeview de resultados
            for item in self.resultados_tree.get_children():
                self.resultados_tree.delete(item)
            
            # Agregar resultados
            for prestamo in resultados:
                self.resultados_tree.insert('', 'end', values=(
                    prestamo['id'],
                    prestamo['usuario'],
                    prestamo['prestamista'],
                    prestamo['equipo'],
                    prestamo.get('estado_equipo', 'Completo'),
                    prestamo['fecha_prestamo'],
                    prestamo.get('fecha_entrega', 'Pendiente'),
                    prestamo.get('quien_recibe', ''),
                    prestamo['estado']
                ))
            
            messagebox.showinfo("Búsqueda", f"Se encontraron {len(resultados)} resultado(s)")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error en la búsqueda: {str(e)}")
    
    def exportar_excel(self):
        """Exportar todos los datos a Excel"""
        try:
            # Crear DataFrame con préstamos
            df_prestamos = pd.DataFrame(self.prestamos)
            
            # Crear DataFrame con equipos
            df_equipos = pd.DataFrame(self.equipos)
            
            # Crear DataFrame con usuarios
            df_usuarios = pd.DataFrame(self.usuarios)
            
            # Crear DataFrame con prestamistas
            df_prestamistas = pd.DataFrame(self.prestamistas)
            
            # Solicitar archivo de destino
            archivo = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Guardar archivo Excel"
            )
            
            if archivo:
                # Crear archivo Excel con múltiples hojas
                with pd.ExcelWriter(archivo, engine='openpyxl') as writer:
                    df_prestamos.to_excel(writer, sheet_name='Préstamos', index=False)
                    df_equipos.to_excel(writer, sheet_name='Equipos', index=False)
                    df_usuarios.to_excel(writer, sheet_name='Usuarios', index=False)
                    df_prestamistas.to_excel(writer, sheet_name='Prestamistas', index=False)
                
                messagebox.showinfo("Éxito", f"Archivo exportado correctamente: {archivo}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar: {str(e)}")
    
    def exportar_prestamos_activos(self):
        """Exportar solo préstamos activos a Excel"""
        try:
            # Filtrar préstamos activos
            prestamos_activos = [p for p in self.prestamos if p['estado'] == 'Prestado']
            
            if not prestamos_activos:
                messagebox.showinfo("Información", "No hay préstamos activos para exportar")
                return
            
            # Crear DataFrame
            df = pd.DataFrame(prestamos_activos)
            
            # Solicitar archivo de destino
            archivo = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Guardar préstamos activos"
            )
            
            if archivo:
                df.to_excel(archivo, index=False)
                messagebox.showinfo("Éxito", f"Préstamos activos exportados: {archivo}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar: {str(e)}")
    
    def ejecutar(self):
        """Ejecutar la aplicación"""
        self.root.mainloop()

if __name__ == "__main__":
    app = SistemaPrestamos()
    app.ejecutar()
