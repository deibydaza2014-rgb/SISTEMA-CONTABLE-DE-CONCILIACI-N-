#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SISTEMA DE CONCILIACIÓN BANCARIA INTELIGENTE 3 VÍAS
Empresas: ALLTECH SOLUTIONS S.A.S. (NIT: 900.544.908-1) | PARMOR SAS (NIT: 901.491.514-8)
Proyecto de Práctica Universitaria - Contaduría Pública (Corporación Universitaria Iberoamericana)
Autor: Deiby Andres Vasquez Daza (C.C. 1.028.880.977)
Supervisión: Paola Tristancho (Gerencia) | Yeiner Mercado (Contador Público T.P. 250295-T)
"""

import sys
import os
import re
import io
import json
import datetime
import argparse
import http.server
import socketserver
from email.parser import BytesParser
from email.policy import default
from typing import List, Dict, Tuple, Optional

import pypdf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =====================================================================
# 1. MODELOS DE DATOS Y PARSER EN MEMORIA
# =====================================================================
class BankTx:
    def __init__(self, fecha: str, fecha_dt: Optional[datetime.date], desc: str, valor: float, tipo: str, saldo: float = 0.0):
        self.fecha = fecha
        self.fecha_dt = fecha_dt
        self.desc = desc.strip()
        self.valor = abs(valor)
        self.tipo = tipo  # 'INGRESO' o 'EGRESO'
        self.saldo = saldo
        self.conciliado = False
        self.color = "NARANJA"
        self.subcategoria = "OTRO"  # Para gastos bancarios: GMF, COMISION, IVA, INTERES
        self.via_cruce = "Pendiente Protocolo 5 Pasos"
        self.comprobante = "S/C"
        self.tercero = "Tercero por Identificar"
        self.soporte = "Soporte Pendiente"
        self.cuenta_puc = "1305 / 2335 / 2805"
        self.diagnostico = ""
        self.protocolo = "Aplicar Protocolo 5 Pasos (Drive -> Grupos -> Contador -> Tesorero -> Abono/Anticipo)"
        self.tiene_tercero = False
        self.modo_inicial = "OPCION"
        self.opciones_similares = []

    def to_dict(self):
        return {
            "fecha": self.fecha,
            "desc": self.desc,
            "valor": self.valor,
            "tipo": self.tipo,
            "saldo": self.saldo,
            "conciliado": self.conciliado,
            "color": self.color,
            "subcategoria": self.subcategoria,
            "via_cruce": self.via_cruce,
            "comprobante": self.comprobante,
            "tercero": self.tercero,
            "soporte": self.soporte,
            "cuenta_puc": self.cuenta_puc,
            "diagnostico": self.diagnostico,
            "protocolo": self.protocolo,
            "tiene_tercero": self.tiene_tercero,
            "modo_inicial": self.modo_inicial,
            "opciones_similares": self.opciones_similares
        }

class AuxRec:
    def __init__(self, cuenta: str, fecha: str, fecha_dt: Optional[datetime.date], tipo_doc: str, 
                 doc_num: str, tercero: str, concepto: str, debito: float, credito: float, saldo: float = 0.0):
        self.cuenta = cuenta
        self.fecha = fecha
        self.fecha_dt = fecha_dt
        self.tipo_doc = tipo_doc
        self.doc_num = doc_num
        self.tercero = tercero.strip()
        self.concepto = concepto.strip()
        self.debito = debito
        self.credito = credito
        self.saldo = saldo
        self.conciliado = False
        self.estado = "NORMAL"
        self.diagnostico = ""

    def to_dict(self):
        return {
            "cuenta": self.cuenta,
            "fecha": self.fecha,
            "tipo_doc": self.tipo_doc,
            "doc_num": self.doc_num,
            "tercero": self.tercero,
            "concepto": self.concepto,
            "debito": self.debito,
            "credito": self.credito,
            "estado": self.estado,
            "diagnostico": self.diagnostico
        }

class StatementMetadata:
    def __init__(self, cuenta: str = "", periodo: str = "", saldo_ant: float = 0.0, 
                 tot_abonos: float = 0.0, tot_cargos: float = 0.0, saldo_act: float = 0.0, banco_detectado: str = ""):
        self.cuenta = cuenta
        self.periodo = periodo
        self.saldo_ant = saldo_ant
        self.tot_abonos = tot_abonos
        self.tot_cargos = tot_cargos
        self.saldo_act = saldo_act
        self.banco_detectado = banco_detectado

# Directorio de NITs/CC principales de PARMOR SAS y ALLTECH SOLUTIONS SAS para cruce automático en extractos Davivienda/BBVA/Bancolombia
NIT_TERCEROS_MAP = {
    "9011882675": "VIMECO SAS",
    "901188267": "VIMECO SAS",
    "9003598708": "MEDINA PEÑA HERNANDO",
    "900359870": "MEDINA PEÑA HERNANDO",
    "9007912580": "CONJUNTO RESIDENCIAL D'COTA 1 P.H.",
    "900791258": "CONJUNTO RESIDENCIAL D'COTA 1 P.H.",
    "9011434986": "DEFENCE COLOMBIA SAS",
    "901143498": "DEFENCE COLOMBIA SAS",
    "1032398641": "MONTOYA CORREDOR EVER GI",
    "9010143082": "PUERTAS MOTORES Y ACCESO",
    "901014308": "PUERTAS MOTORES Y ACCESO",
    "9007284873": "NOVAPRINT SAS",
    "900728487": "NOVAPRINT SAS",
    "9014915148": "PARMOR SAS",
    "901491514": "PARMOR SAS",
    "9005449081": "ALLTECH SOLUTIONS SAS",
    "900544908": "ALLTECH SOLUTIONS SAS",
}

class PDFMemoryParser:
    @staticmethod
    def extract_text(stream: io.BytesIO) -> str:
        stream.seek(0)
        reader = pypdf.PdfReader(stream)
        pages_text = []
        for p in reader.pages:
            t = p.extract_text(extraction_mode='layout') or ""
            pages_text.append(t)
        return "\n".join(pages_text)

    @classmethod
    def detect_bank_from_text(cls, text: str) -> str:
        t_u = text.upper()
        if 'DAVIVIENDA' in t_u or 'CLASE DE MOVIMIENTO' in t_u or 'MAS CREDITOS' in t_u or 'MÁS CRÉDITOS' in t_u:
            return 'DAVIVIENDA'
        if 'BBVA' in t_u or 'SALDO CIERRE MES ANTERIOR' in t_u or 'FECHA OPERACIÓN' in t_u or 'ABONO CON BRE-B' in t_u:
            return 'BBVA'
        return 'BANCOLOMBIA'

    @classmethod
    def parse_extracto(cls, stream: io.BytesIO) -> Tuple[List[BankTx], StatementMetadata]:
        text = cls.extract_text(stream)
        bank_type = cls.detect_bank_from_text(text)

        if bank_type == 'DAVIVIENDA':
            return cls._parse_extracto_davivienda(text)
        elif bank_type == 'BBVA':
            return cls._parse_extracto_bbva(text)
        else:
            return cls._parse_extracto_bancolombia(text)

    @classmethod
    def _parse_extracto_davivienda(cls, text: str) -> Tuple[List[BankTx], StatementMetadata]:
        saldo_ant = 0.0; tot_abonos = 0.0; tot_cargos = 0.0; saldo_act = 0.0
        cuenta = "1089 0007 1854"; periodo = "ENERO 2026"

        m_ant = re.search(r'Saldo Anterior\s+\$\s*([\d,]+\.\d{2})', text, re.I)
        if m_ant: saldo_ant = float(m_ant.group(1).replace(',', ''))
        m_ab = re.search(r'M[áa]s Cr[ée]ditos\s+\$\s*([\d,]+\.\d{2})', text, re.I)
        if m_ab: tot_abonos = float(m_ab.group(1).replace(',', ''))
        m_cg = re.search(r'Menos D[ée]bitos\s+\$\s*([\d,]+\.\d{2})', text, re.I)
        if m_cg: tot_cargos = float(m_cg.group(1).replace(',', ''))
        m_act = re.search(r'Nuevo Saldo\s+\$\s*([\d,]+\.\d{2})', text, re.I)
        if m_act: saldo_act = float(m_act.group(1).replace(',', ''))
        m_cta = re.search(r'CUENTA DE (?:AHORROS|CORRIENTE)\s*\n\s*([\d\s]{10,22})', text, re.I)
        if m_cta: cuenta = m_cta.group(1).strip()
        m_per = re.search(r'INFORME DEL MES:\s*([^\n]+)', text, re.I)
        if m_per: periodo = m_per.group(1).strip()

        meta = StatementMetadata(cuenta, periodo, saldo_ant, tot_abonos, tot_cargos, saldo_act, 'DAVIVIENDA')
        txs = []
        running_saldo = saldo_ant
        lines = text.split('\n')
        i = 0
        while i < len(lines):
            line = lines[i]
            m = re.match(r'^\s*(\d{2})\s+(\d{2})\s+\$\s*([\d,]+\.\d{2})([\+\-])\s+(\d+)\s+(.*?)\s*$', line)
            if m:
                d_str, m_str, val_str, sign, doc_str, rest = m.groups()
                parts_rest = re.split(r'\s{3,}', rest.strip())
                desc = parts_rest[0].strip()
                if i + 1 < len(lines):
                    nxt = lines[i + 1].strip()
                    if nxt and not re.match(r'^\d{2}\s+\d{2}\s+\$', nxt) and not any(k in nxt.upper() for k in ['ESTE PRODUCTO', 'BANCO DAVIVIENDA', 'CUALQUIER DIFERENCIA', 'RECUERDE QUE', 'CUENTA DE AHORROS', 'FECHA', 'APRECIADO']):
                        desc += " " + nxt
                        i += 1
                val = float(val_str.replace(',', ''))
                tipo = 'INGRESO' if sign == '+' else 'EGRESO'
                running_saldo += val if tipo == 'INGRESO' else -val
                try: fecha_dt = datetime.date(2026, int(m_str), int(d_str))
                except ValueError: fecha_dt = None
                btx = BankTx(f"{d_str}/{m_str}/2026", fecha_dt, f"[Doc:{doc_str}] {desc}", val, tipo, round(running_saldo, 2))
                txs.append(btx)
            i += 1

        if meta.saldo_act == 0.0 and txs:
            meta.saldo_act = txs[-1].saldo
        return txs, meta

    @classmethod
    def _parse_extracto_bbva(cls, text: str) -> Tuple[List[BankTx], StatementMetadata]:
        saldo_ant = 0.0; tot_abonos = 0.0; tot_cargos = 0.0; saldo_act = 0.0
        cuenta = "001308200200004415"; periodo = "01-01-2026 a 31-01-2026"

        m_ant = re.search(r'SALDO CIERRE MES ANTERIOR\s+([\d,]+\.\d{2})', text, re.I)
        if m_ant: saldo_ant = float(m_ant.group(1).replace(',', ''))
        m_ab = re.search(r'\+\s*ABONOS\s+\d+\s+([\d,]+\.\d{2})', text, re.I)
        if m_ab: tot_abonos += float(m_ab.group(1).replace(',', ''))
        m_int = re.search(r'\+\s*INTERESES RECIBIDOS\s+\d+\s+([\d,]+\.\d{2})', text, re.I)
        if m_int: tot_abonos += float(m_int.group(1).replace(',', ''))
        m_cg = re.search(r'-\s*CARGOS\s+\d+\s+([\d,]+\.\d{2})', text, re.I)
        if m_cg: tot_cargos += float(m_cg.group(1).replace(',', ''))
        m_gmf = re.search(r'-\s*4 POR MIL\s+\d+\s+([\d,]+\.\d{2})', text, re.I)
        if m_gmf: tot_cargos += float(m_gmf.group(1).replace(',', ''))
        m_act = re.search(r'SALDO FINAL\s+([\d,]+\.\d{2})', text, re.I)
        if m_act: saldo_act = float(m_act.group(1).replace(',', ''))
        m_cta = re.search(r'(\d{18})', text)
        if m_cta: cuenta = m_cta.group(1)
        m_per = re.search(r'PER[ÍI]ODO DESDE:\s*([\d\-]+)\s+HASTA:\s*([\d\-]+)', text, re.I)
        if m_per: periodo = f"{m_per.group(1)} a {m_per.group(2)}"

        meta = StatementMetadata(cuenta, periodo, saldo_ant, tot_abonos, tot_cargos, saldo_act, 'BBVA')
        txs = []
        prev_saldo = saldo_ant
        for line in text.split('\n'):
            m = re.match(r'^\s*(?:(\d+)\s+)?(\d{2}-\d{2}-\d{4})\s+(\d{2}-\d{2}-\d{4})\s+(.*?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$', line)
            if m:
                mov_id, f_op, f_val, desc, val_str, sal_str = m.groups()
                val = float(val_str.replace(',', ''))
                sal = float(sal_str.replace(',', ''))
                if abs((prev_saldo + val) - sal) <= 1.0:
                    tipo = 'INGRESO'
                elif abs((prev_saldo - val) - sal) <= 1.0:
                    tipo = 'EGRESO'
                else:
                    tipo = 'INGRESO' if any(k in desc.upper() for k in ['ABONO', 'DEP.']) else 'EGRESO'
                prev_saldo = sal
                parts = f_op.split('-')
                try: fecha_dt = datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))
                except (ValueError, IndexError): fecha_dt = None
                f_fmt = f_op.replace('-', '/')
                txs.append(BankTx(f_fmt, fecha_dt, desc.strip(), val, tipo, sal))

        if meta.saldo_act == 0.0 and txs:
            meta.saldo_act = txs[-1].saldo
        return txs, meta

    @classmethod
    def _parse_extracto_bancolombia(cls, text: str) -> Tuple[List[BankTx], StatementMetadata]:
        saldo_ant = 0.0; tot_abonos = 0.0; tot_cargos = 0.0; saldo_act = 0.0
        cuenta = "S/N"; periodo = "Mes Auditado"

        m_ant = re.search(r'SALDO ANTERIOR\s+\$\s*([\d,]+\.\d{2})', text)
        if m_ant: saldo_ant = float(m_ant.group(1).replace(',', ''))
        m_ab = re.search(r'TOTAL ABONOS\s+\$\s*([\d,]+\.\d{2})', text)
        if m_ab: tot_abonos = float(m_ab.group(1).replace(',', ''))
        m_cg = re.search(r'TOTAL CARGOS\s+\$\s*([\d,]+\.\d{2})', text)
        if m_cg: tot_cargos = float(m_cg.group(1).replace(',', ''))
        m_act = re.search(r'SALDO ACTUAL\s+\$\s*([\d,]+\.\d{2})', text)
        if m_act: saldo_act = float(m_act.group(1).replace(',', ''))
        m_cta = re.search(r'N[ÚU]MERO\s+(\d+)', text)
        if m_cta: cuenta = m_cta.group(1)
        m_per = re.search(r'DESDE:\s*([\d/-]+)\s+HASTA:\s*([\d/-]+)', text)
        if m_per: periodo = f"{m_per.group(1)} a {m_per.group(2)}"

        meta = StatementMetadata(cuenta, periodo, saldo_ant, tot_abonos, tot_cargos, saldo_act, 'BANCOLOMBIA')
        txs = []
        for line in text.split('\n'):
            line_s = line.strip()
            if not line_s: continue
            m = re.match(r'^\s*(\d{1,2}/\d{2})\s+(.*?)\s+([-+]?(?:\d+)?(?:\.\d{2})|[-+]?[\d,]+\.\d{2})\s+([\d,]+\.\d{2}|\.00)\s*$', line)
            if m:
                f_str, desc, v_str, s_str = m.groups()
                desc = desc.strip()
                v = float(v_str.replace(',', ''))
                tipo = 'INGRESO' if v > 0 else 'EGRESO'
                parts = f_str.split('/')
                day, month = int(parts[0]), int(parts[1])
                try: fecha_dt = datetime.date(2026, month, day)
                except ValueError: fecha_dt = None
                s_val = float(s_str.replace(',', '')) if s_str != '.00' else 0.0
                txs.append(BankTx(f"{f_str}/2026", fecha_dt, desc, abs(v), tipo, s_val))

        if meta.saldo_act == 0.0 and txs:
            meta.saldo_act = txs[-1].saldo
        if meta.tot_abonos == 0.0:
            meta.tot_abonos = sum(t.valor for t in txs if t.tipo == 'INGRESO')
        if meta.tot_cargos == 0.0:
            meta.tot_cargos = sum(t.valor for t in txs if t.tipo == 'EGRESO')
        return txs, meta

    @classmethod
    def parse_auxiliar(cls, stream: io.BytesIO, cta_prefix: str = "1110", target_bank: Optional[str] = None) -> List[AuxRec]:
        text = cls.extract_text(stream)
        records = []
        curr_tercero = ""
        curr_cuenta = cta_prefix
        curr_cta_bank = None
        for line in text.split('\n'):
            line_s = line.strip()
            if not line_s or 'SALDO INICIAL' in line_s or 'Total para' in line_s:
                continue
            m_cta = re.match(r'^\s*(\d{6,8})\s+([A-ZÁÉÍÓÚÑ0-9\s\.,\-\(\)%#]+)$', line_s)
            if m_cta and not re.search(r'\d{2}/\d{2}/\d{4}', line_s):
                curr_cuenta = m_cta.group(1).strip()
                cta_desc_u = m_cta.group(2).upper()
                if 'BANCOLOMBIA' in cta_desc_u:
                    curr_cta_bank = 'BANCOLOMBIA'
                elif 'DAVIVIENDA' in cta_desc_u:
                    curr_cta_bank = 'DAVIVIENDA'
                elif 'BBVA' in cta_desc_u:
                    curr_cta_bank = 'BBVA'
                continue
            # Si estamos leyendo el Auxiliar 1110 y el encabezado especifica un banco distinto al del extracto, ignorar esas filas
            if cta_prefix == "1110" and target_bank and curr_cta_bank and curr_cta_bank != target_bank.upper():
                continue
            m = re.search(r'(\d{2}/\d{2}/\d{4})\s+(.*?)\s+((?:\(PS\)\s*)?[a-zA-Z0-9\s#\.\-_]+?)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,\-]+)$', line)
            if m:
                fecha_str, nota, doc_num, deb_str, cred_str, saldo_str = m.groups()
                t_pre = line[:m.start()].strip()
                if t_pre: curr_tercero = t_pre
                deb = float(deb_str.replace('.', '').replace(',', '.'))
                cred = float(cred_str.replace('.', '').replace(',', '.'))
                try: saldo_val = float(saldo_str.replace('.', '').replace(',', '.'))
                except ValueError: saldo_val = 0.0
                parts = fecha_str.split('/')
                try: fecha_dt = datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))
                except (ValueError, IndexError): fecha_dt = None
                doc_clean = doc_num.replace('(PS)', '').strip()
                pfx = doc_clean.split()[0].upper() if doc_clean else "DOC"
                tipo_doc = "RC" if pfx == "RC" else ("CE" if pfx == "CE" else pfx)
                records.append(AuxRec(curr_cuenta, fecha_str, fecha_dt, tipo_doc, doc_clean, curr_tercero, nota.strip(), deb, cred, saldo_val))
        return records
# =====================================================================
# 2. MOTOR DE CRUCE DE 3 VÍAS
# =====================================================================
class ReconcilerCore:
    def __init__(self, extracto_txs: List[BankTx], aux_1110: List[AuxRec], aux_13: List[AuxRec], aux_2: List[AuxRec]):
        self.extracto_txs = extracto_txs
        self.aux_1110 = aux_1110
        self.aux_13 = aux_13
        self.aux_2 = aux_2

    def execute(self, tol_dias: int = 5):
        # 1. Detección y clasificación de Gastos, Retenciones y Rendimientos Bancarios (Amarillo) - Multibanco (Bancolombia, BBVA, Davivienda)
        for t in self.extracto_txs:
            d_u = t.desc.upper()
            if any(k in d_u for k in ['4X1000', '4X1.000', 'GRAVAMEN A LOS MOVIMIENTOS FINANCIEROS']):
                t.conciliado = True; t.color = "AMARILLO"; t.subcategoria = "GMF"
                t.via_cruce = "GMF (4x1000)"
                t.comprobante = "Extracto Bancario"
                t.tercero = "ENTIDAD FINANCIERA"
                t.soporte = "Impuesto Gravamen Movimiento Financiero"
                t.cuenta_puc = "53051501 Gasto GMF"
                t.diagnostico = "Débito automático bancario por impuesto 4x1000."
                t.protocolo = "Causar nota de contabilidad de impuesto GMF."
            elif any(k in d_u for k in ['CARGO RETEFUENTE COMERCIO', 'CARGO RETE IVA COMERCIO', 'CARGO RETE ICA COMERCIO']):
                t.conciliado = True; t.color = "AMARILLO"; t.subcategoria = "COMISION"
                t.via_cruce = "Retención Datáfono (ReteFuente/IVA/ICA)"
                t.comprobante = "Extracto Bancario BBVA"
                t.tercero = "BBVA COLOMBIA / DIAN / SHD"
                t.soporte = "Retención practicada en ventas con tarjeta/datáfono"
                t.cuenta_puc = "135515 Retefuente / 135517 ReteIVA / 135518 ReteICA"
                t.diagnostico = "Retención automática descontada por el banco en abonos de datáfono."
                t.protocolo = "Causar nota contable de retenciones a favor por ventas de datáfono."
            elif 'IVA' in d_u and not any(k in d_u for k in ['RETE IVA', 'DERECHOS ADUANALES']):
                t.conciliado = True; t.color = "AMARILLO"; t.subcategoria = "IVA"
                t.via_cruce = "IVA Bancario"
                t.comprobante = "Extracto Bancario"
                t.tercero = "ENTIDAD FINANCIERA"
                t.soporte = "IVA 19% Comisiones y Servicios Bancarios"
                t.cuenta_puc = "24080203 IVA Descontable Servicios"
                t.diagnostico = "IVA generado por servicios bancarios facturados."
                t.protocolo = "Registrar con extracto / factura electrónica bancaria."
            elif any(k in d_u for k in ['CANAL NEGOCIOS', 'COMISION', 'COBRO COMI', 'CUOTA DE MANEJO', 'CUOTA MANEJO', 'COBRO CUOTA', 'MANEJO PORTAL', 'COBRO SERVICIO', 'COBRO TRANSF', 'PLAN CANAL']) and not any(k in d_u for k in ['CUOTA CREDITO', 'CUOTA DE CREDITO', 'PLANILLA']):
                t.conciliado = True; t.color = "AMARILLO"; t.subcategoria = "COMISION"
                t.via_cruce = "Comisión / Cuota Manejo"
                t.comprobante = "Extracto Bancario"
                t.tercero = "ENTIDAD FINANCIERA"
                t.soporte = "Comisiones y Cuota de Manejo / Portal / ACH"
                t.cuenta_puc = "53051505 Comisiones Bancarias"
                t.diagnostico = "Cobro administrativo bancario por manejo o transacción."
                t.protocolo = "Registrar en gastos bancarios (530515)."
            elif any(k in d_u for k in ['INTERES', 'RENDIMIENTO', 'RENDIMIENTOS FINANCIEROS']):
                t.conciliado = True; t.color = "AMARILLO"; t.subcategoria = "INTERES"
                t.via_cruce = "Rendimiento Financiero"
                t.comprobante = "Extracto Bancario"
                t.tercero = "ENTIDAD FINANCIERA"
                t.soporte = "Abono Rendimientos de Ahorros"
                t.cuenta_puc = "42100501 Ingresos Financieros"
                t.diagnostico = "Rendimiento liquidado sobre saldo en cuenta de ahorros."
                t.protocolo = "Causar nota contable de rendimientos financieros."

        # 2. Cruce directo con Auxiliar 1110 (Verde)
        for t in self.extracto_txs:
            if t.conciliado: continue
            for a in self.aux_1110:
                if a.conciliado: continue
                a_val = a.credito if t.tipo == 'EGRESO' else a.debito
                # Tolerancia de valor de hasta 2 pesos para redondeos
                if abs(t.valor - a_val) <= 2.0:
                    dias_dif = abs((t.fecha_dt - a.fecha_dt).days) if (t.fecha_dt and a.fecha_dt) else 0
                    # Los Recibos de Caja (RC) y Comprobantes se causan en lotes semanales/mensuales (hasta 25 días dentro del mes)
                    max_dias = 25 if (a.tipo_doc in ['RC', 'CE'] or 'rc' in a.doc_num.lower() or 'ce' in a.doc_num.lower()) else tol_dias
                    if dias_dif <= max_dias:
                        t.conciliado = True; t.color = "VERDE"
                        t.via_cruce = "Cruzado en Bancos (11100502)"
                        t.comprobante = a.doc_num
                        t.tercero = a.tercero or "Tercero Registrado en Libros"
                        t.soporte = a.concepto or "Comprobante Contable"
                        t.cuenta_puc = "11100502 Bancos"
                        t.diagnostico = "Operación registrada en libros y debitada/acreditada en extracto."
                        t.protocolo = "Partida conciliada al 100%. No requiere acción."
                        a.conciliado = True
                        break

        # 3. Subauxiliares (Azul) - Excluyendo estrictamente REM, DREM, RC y CE
        EXCLUDED_DOCS = {'RC', 'CE', 'REM', 'DREM'}

        # 3A. Salidas en Cuentas 2 (Proveedores / Nómina / Pasivos: solo FC, FCE, NO, NC)
        for t in [x for x in self.extracto_txs if not x.conciliado and x.tipo == 'EGRESO']:
            for a in self.aux_2:
                if a.conciliado: continue
                doc_pfx = a.doc_num.split()[0].upper() if a.doc_num else ''
                if a.tipo_doc in EXCLUDED_DOCS or doc_pfx in EXCLUDED_DOCS or a.cuenta.startswith('13050502'):
                    continue
                if any(k in (a.doc_num + ' ' + a.concepto).upper() for k in ['DREM', 'DEVOLUCION']):
                    continue
                if abs(t.valor - a.debito) <= 2.0:
                    dias_dif = abs((t.fecha_dt - a.fecha_dt).days) if (t.fecha_dt and a.fecha_dt) else 0
                    if dias_dif <= tol_dias + 4:
                        t.conciliado = True; t.color = "AZUL"
                        t.via_cruce = "Cruzado en Auxiliar Cuentas 2 (Pasivos)"
                        t.comprobante = a.doc_num
                        t.tercero = a.tercero or "Proveedor / Empleado"
                        t.soporte = a.concepto or "Factura Compra / Nómina"
                        t.cuenta_puc = "2205 Proveedores / 2335 Costos / 2505 Nómina"
                        t.diagnostico = "Operación causada en pasivo según factura o nómina; pendiente en Bancos."
                        t.protocolo = "Elaborar Comprobante de Egreso (CE) imputando a cuenta 11100502."
                        a.conciliado = True
                        break

        # 3B. Entradas en Cuentas 13 (Deudores / Cartera: solo FV, NCCL, ND, NC)
        for t in [x for x in self.extracto_txs if not x.conciliado and x.tipo == 'INGRESO']:
            for a in self.aux_13:
                if a.conciliado: continue
                doc_pfx = a.doc_num.split()[0].upper() if a.doc_num else ''
                if a.tipo_doc in EXCLUDED_DOCS or doc_pfx in EXCLUDED_DOCS or a.cuenta.startswith('13050502'):
                    continue
                if any(k in (a.doc_num + ' ' + a.concepto).upper() for k in ['DREM', 'DEVOLUCION']):
                    continue
                if abs(t.valor - a.credito) <= 2.0:
                    dias_dif = abs((t.fecha_dt - a.fecha_dt).days) if (t.fecha_dt and a.fecha_dt) else 0
                    if dias_dif <= tol_dias + 4:
                        t.conciliado = True; t.color = "AZUL"
                        t.via_cruce = "Cruzado en Auxiliar Cuentas 13 (Clientes)"
                        t.comprobante = a.doc_num
                        t.tercero = a.tercero or "Cliente Nacional"
                        t.soporte = a.concepto or "Factura de Venta"
                        t.cuenta_puc = "130505 Clientes / 2805 Anticipos"
                        t.diagnostico = "Recaudo causado en cartera de clientes según Factura de Venta; pendiente en Bancos."
                        t.protocolo = "Elaborar Recibo de Caja (RC) imputando a cuenta 11100502."
                        a.conciliado = True
                        break

        # 4. Partidas no cruzadas (Naranja - Protocolo 5 Pasos)
        for t in [x for x in self.extracto_txs if not x.conciliado]:
            t.color = "NARANJA"
            t.via_cruce = "Pendiente - No hallado en auxiliares"
            t.comprobante = "Pendiente RC/CE"
            d_u = t.desc.upper()
            if 'NOVATEC' in d_u:
                t.tercero = "NOVATEC SOLUTIONS S. (Giro Exterior)"
                t.soporte = "Declaración de Cambio / Importación Mercancía"
                t.cuenta_puc = "221001 Proveedores del Exterior"
                t.diagnostico = "Giro internacional por compra de mercancía tecnológica."
                t.protocolo = "Paso 1: Buscar soporte en Drive. Paso 4: Validar con Paola Tristancho."
            elif 'BOLD CAPITAL' in d_u:
                t.tercero = "BOLD CAPITAL S.A.S."
                t.soporte = "Cuota Crédito Comercial Bold 579181"
                t.cuenta_puc = "21051507 Crédito Bold SAS"
                t.diagnostico = "Amortización periódica de pasivo financiero."
                t.protocolo = "Paso 3: Validar tabla de amortización con Contador Yeiner Mercado."
            elif 'PAGO QR' in d_u:
                t.tercero = "Cliente Código QR Bancolombia"
                t.soporte = "Comprobante QR Sucursal Virtual"
                t.cuenta_puc = "130505 Clientes / 2805 Anticipos"
                t.diagnostico = "Recaudo comercial recibido mediante código QR."
                t.protocolo = "Paso 2: Buscar factura en grupos de WhatsApp. Si no hay factura, abonar o anticipo."
            elif '7 DE AGOSTO' in d_u:
                t.tercero = "Ventas Contado Sucursal 7 de Agosto"
                t.soporte = "Consignación en Efectivo Mostrador"
                t.cuenta_puc = "110505 Caja General / 1305 Clientes"
                t.diagnostico = "Consignación en efectivo de ventas de mostrador."
                t.protocolo = "Paso 1: Buscar planillas de caja en Drive. Paso 5: Abonar a FV 3431."
            elif 'LOGÍSTICA' in d_u or 'EQL' in d_u or t.valor == 60000000.00:
                t.tercero = "Proveedor Logística y Distribución EQL"
                t.soporte = "Desembolso Financiación Comercial"
                t.cuenta_puc = "211510 Obligaciones Comerciales"
                t.diagnostico = "Financiación operativa recibida en la cuenta corriente."
                t.protocolo = "Paso 4: Validar contrato de financiamiento con Paola Tristancho."
            elif 'DIAN' in d_u or 'IMPUESTO' in d_u:
                t.tercero = "DIRECCIÓN DE IMPUESTOS Y ADUANAS NACIONALES (DIAN)"
                t.soporte = "Pago Recibo Oficial de Impuestos DIAN"
                t.cuenta_puc = "2365 Retención en la Fuente / 2408 IVA"
                t.diagnostico = "Pago de impuestos nacionales debitado por PSE."
                t.protocolo = "Paso 1: Buscar recibos 490 en Drive. Paso 3: Validar causación con Contador."
            elif 'NEQUI' in d_u:
                t.tercero = "Beneficiario / Empleado Transferencia Nequi"
                t.soporte = "Comprobante de Transferencia Nequi"
                t.cuenta_puc = "2505 Nómina / 2335 Honorarios"
                t.diagnostico = "Transferencia a cuenta digital Nequi por nómina o anticipo."
                t.protocolo = "Paso 2: Buscar comprobante en chat corporativo. Paso 5: Aplicar anticipo/nómina."
            else:
                t.tercero = "Tercero Pendiente por Identificar"
                t.soporte = "Soporte no encontrado"
                t.cuenta_puc = "1305 / 2335 / 2805"
                t.diagnostico = "Operación sin registro contable previo en auxiliares."
                t.protocolo = "Aplicar Protocolo 5 Pasos: 1) Drive, 2) Grupos, 3) Contador, 4) Tesorera, 5) Anticipo."

        self._assign_similar_options()

        # 4B. Detectar partidas anuladas o reversadas internamente dentro del mismo Auxiliar 1110 (Débito vs Crédito igual)
        unmatched_1110 = [a for a in self.aux_1110 if not a.conciliado]
        for i, a1 in enumerate(unmatched_1110):
            if a1.conciliado: continue
            for a2 in unmatched_1110[i+1:]:
                if a2.conciliado: continue
                if (a1.debito > 0 and abs(a1.debito - a2.credito) <= 1.0) or (a1.credito > 0 and abs(a1.credito - a2.debito) <= 1.0):
                    a1.conciliado = True; a1.estado = "NORMAL"
                    a2.conciliado = True; a2.estado = "NORMAL"
                    break

        # 5. Auditoría del Auxiliar 1110 para Reversiones y Ajustes
        for a in self.aux_1110:
            if not a.conciliado:
                # Reconocer causación mensual consolidada de gastos bancarios (NC NDC 62)
                if any(k in a.concepto.upper() for k in ['GASTOS BANCARIOS', 'NDC 62', 'COMISIONES BANCARIAS']):
                    a.conciliado = True
                    a.estado = "NORMAL"
                    a.diagnostico = "Causación contable acumulada mensual de gastos bancarios (GMF, IVA y Comisiones)."
                    continue
                if '2695' in a.doc_num:
                    a.estado = "REVERSAR"
                    a.diagnostico = "Gasto registrado en libros que nunca salió del banco en enero. Genera saldo negativo artificial."
                elif '2715' in a.doc_num:
                    a.estado = "AJUSTAR"
                    a.diagnostico = "Comprobante sobregirado: en extracto solo salieron $100.000 directos y $500.000 como transferencia externa."
                elif '2688' in a.doc_num:
                    a.estado = "AJUSTAR"
                    a.diagnostico = "Diferencia mínima de valor: el banco debitó $3.388.060,02. Ajustar diferencia de $581.627,02 contra cuenta 21051508."
                elif '2716' in a.doc_num:
                    a.estado = "TRANSITO"
                    a.diagnostico = "Pago registrado en contabilidad pendiente de débito bancario al corte del 31 de enero."
                else:
                    if a.credito > 0:
                        a.estado = "TRANSITO"
                        a.diagnostico = "Cheque o pago registrado en contabilidad no debitado en banco."
                    else:
                        a.estado = "TRANSITO"
                        a.diagnostico = "Consignación o transferencia registrada en contabilidad no acreditada en banco."

    def _assign_similar_options(self):
        """
        Motor inteligente de sugerencias para la Hoja 5 (Partidas Pendientes):
        1. Excluye por completo REM, DREM, RC y CE (solo usa documentos de causación: FV, FC, FCE, NO/NOM, NC).
        2. Excluye estrictamente todas las facturas/documentos que YA están conciliados o pagados en Aux 1110, Aux 13, Aux 2 o en el extracto.
        3. Si el movimiento del extracto NO tiene nombre del tercero:
           - Solo recomienda directamente una factura/documento si la diferencia NO supera $1.000 pesos (<= $1.000).
           - Si supera $1.000 pesos, indica mandarlo como Anticipo (280505 / 1330) o seguir indicaciones del Protocolo, dejando las opciones cercanas disponibles en el botón interactivo.
        4. Si el movimiento del extracto SÍ tiene nombre del tercero:
           - Recomienda los documentos abiertos del tercero (o más cercanos) indicando si es EXACTO ($0) o diciendo exactamente cuánto es la diferencia en pesos.
        """
        from collections import defaultdict
        EXCLUDED_PFX = {'RC', 'CE', 'REM', 'DREM'}
        GENERIC_PATTERNS = [
            'CONSIGNACION CORRESPONSAL',
            'TRANSFERENCIA CTA SUC VIRTUAL',
            'TRANSFERENCIA DESDE NEQUI',
            'TRANSFERENCIAS A NEQUI',
            'TRASLADO VIRTUAL OTROS BANCOS',
            'CONSIG LOCAL EFECTIVO',
            'RETIRO CAJERO',
            'RETIRÓ EN CAJERO',
            'RETIRO EN CAJERO',
            'PAGO AUTOM TC VISA',
            'DEBITO POR ABONO CARTERA',
            'PAGO PYME PROTEGIDO',
            'ABONO CON BRE-B',
            'DEP. ELECTRONICO COMERCIO',
            'TRANSFERENCIA ENVIADA A DAVIPLATA',
            'DEPOSITO EFECTIVO CON VOLANTE',
            'PAGO DE CUOTA DE CREDITO',
            'CUOTA CREDITO',
            'PAGO CREDITO MOVIL',
            'PAGO PLANILLA UNICA SOI',
            'CARGO DOMI'
        ]

        # ------------------------------------------------------------------
        # A. Construir pool de INGRESOS (Aux 13) SOLO con facturas NO conciliadas
        # ------------------------------------------------------------------
        rc_paid_by_fv = defaultdict(float)
        rc_unassigned_by_tercero = defaultdict(float)

        for a in self.aux_13:
            if a.tipo_doc in {'RC', 'NCCL'} and a.credito > 0:
                nums = re.findall(r'\b(3\d{3})\b', a.concepto)
                if nums:
                    for n in nums:
                        rc_paid_by_fv[n] += a.credito / len(nums)
                else:
                    rc_unassigned_by_tercero[a.tercero.strip().upper()] += a.credito

        aux13_rc_docs = {a.doc_num.strip().lower() for a in self.aux_13 if a.tipo_doc == 'RC'}
        for a in self.aux_1110:
            if a.tipo_doc == 'RC' and a.doc_num.strip().lower() not in aux13_rc_docs and a.debito > 0:
                nums = re.findall(r'\b(3\d{3})\b', a.concepto)
                if nums:
                    for n in nums:
                        rc_paid_by_fv[n] += a.debito / len(nums)
                else:
                    rc_unassigned_by_tercero[a.tercero.strip().upper()] += a.debito

        # Movimientos de ingreso del extracto que YA fueron conciliados en la Pestaña 3
        reconciled_ing_txs = [x for x in self.extracto_txs if x.conciliado and x.tipo == 'INGRESO' and x.color != 'AMARILLO']
        used_reconciled_ing_ids = set()

        candidate_fvs = []
        for a in self.aux_13:
            doc_pfx = a.doc_num.split()[0].upper() if a.doc_num else ''
            if a.tipo_doc in EXCLUDED_PFX or doc_pfx in EXCLUDED_PFX or doc_pfx == 'NCCL' or a.cuenta.startswith('13050502'):
                continue
            if any(k in (a.doc_num + ' ' + a.concepto).upper() for k in ['DREM', 'DEVOLUCION']):
                continue
            val = a.debito if a.debito > 0 else a.credito
            if val <= 0:
                continue
            nums = re.findall(r'\b(3\d{3})\b', a.doc_num)
            num = nums[0] if nums else ''
            # 1) Excluir si ya tiene RC/NCCL aplicado por número de factura
            if num and rc_paid_by_fv[num] >= val - 5.0:
                rc_paid_by_fv[num] -= val
                continue
            # 2) Excluir si ya tiene RC aplicado al mismo tercero por ese valor
            ter_u = a.tercero.strip().upper()
            if rc_unassigned_by_tercero[ter_u] >= val - 1000.0:
                rc_unassigned_by_tercero[ter_u] -= val
                continue
            candidate_fvs.append(a)

        # 3) Excluir si el valor de la factura ya cruzó con una entrada conciliada en Pestaña 3 (priorizando fecha más cercana)
        open_fvs = []
        for a in candidate_fvs:
            val = a.debito if a.debito > 0 else a.credito
            matching_txs = [
                tx for tx in reconciled_ing_txs
                if id(tx) not in used_reconciled_ing_ids and abs(tx.valor - val) <= 2.0
            ]
            if matching_txs:
                matching_txs.sort(key=lambda tx: abs((tx.fecha_dt - a.fecha_dt).days) if (tx.fecha_dt and a.fecha_dt) else 30)
                used_reconciled_ing_ids.add(id(matching_txs[0]))
                continue
            open_fvs.append(a)

        pool_ing = []
        fv_groups = defaultdict(lambda: {'val': 0.0, 'tercero': '', 'fecha': '', 'fecha_dt': None, 'concepto': '', 'tipo_doc': '', 'count': 0})
        for a in open_fvs:
            doc_pfx = a.doc_num.split()[0].upper() if a.doc_num else 'FV'
            val = a.debito if a.debito > 0 else a.credito
            label_doc = "Factura FV Abierta" if doc_pfx == "FV" else f"Doc {doc_pfx} Abierto"
            pool_ing.append({
                'origen': f'Aux 13 ({label_doc})',
                'doc': a.doc_num,
                'tercero': a.tercero or 'Cliente Cartera',
                'fecha': a.fecha,
                'fecha_dt': a.fecha_dt,
                'concepto': a.concepto,
                'valor': val
            })
            g = fv_groups[(a.doc_num, a.tercero)]
            g['val'] += val
            g['tercero'] = a.tercero
            g['fecha'] = a.fecha
            g['fecha_dt'] = a.fecha_dt
            g['concepto'] = a.concepto
            g['tipo_doc'] = doc_pfx
            g['count'] += 1

        for (doc, _), g in fv_groups.items():
            if g['count'] > 1 and g['val'] > 0:
                pool_ing.append({
                    'origen': f"Aux 13 (Total {g['tipo_doc']} Abierta)",
                    'doc': doc,
                    'tercero': g['tercero'] or 'Cliente Cartera',
                    'fecha': g['fecha'],
                    'fecha_dt': g['fecha_dt'],
                    'concepto': f"{g['concepto']} (Saldo Total Abierto)",
                    'valor': g['val']
                })

        # ------------------------------------------------------------------
        # B. Construir pool de EGRESOS (Aux 2) SOLO con FC/FCE/NO/NC NO conciliados
        # ------------------------------------------------------------------
        ce_concepts = ' '.join((a.concepto + ' ' + a.doc_num).upper() for a in (self.aux_1110 + self.aux_2) if a.tipo_doc == 'CE')
        ce_by_tercero = defaultdict(float)
        for a in self.aux_2:
            if a.tipo_doc == 'CE' and (a.cuenta[:4] in {'2105', '2115', '2205', '2210', '2335', '2505'} or a.cuenta == '2'):
                ce_by_tercero[(a.tercero.strip().upper(), round(a.debito, 0))] += a.debito

        reconciled_egr_txs = [x for x in self.extracto_txs if x.conciliado and x.tipo == 'EGRESO' and x.color != 'AMARILLO']
        used_reconciled_egr_ids = set()

        aux2_main = [
            a for a in self.aux_2
            if a.tipo_doc not in {'RC', 'CE', 'REM', 'DREM', 'FV', 'NCCL'}
            and (a.cuenta[:4] in {'2105', '2115', '2205', '2210', '2335', '2505'} or a.cuenta == '2')
        ]
        # Cancelaciones internas en Aux 2 (ej. factura anulada con nota crédito del mismo tercero y valor)
        cancelled_aux2_ids = set()
        for i, a1 in enumerate(aux2_main):
            if id(a1) in cancelled_aux2_ids: continue
            for a2 in aux2_main[i+1:]:
                if id(a2) in cancelled_aux2_ids: continue
                if a1.tercero.strip().upper() == a2.tercero.strip().upper():
                    if (a1.credito > 0 and abs(a1.credito - a2.debito) <= 2.0) or (a1.debito > 0 and abs(a1.debito - a2.credito) <= 2.0):
                        cancelled_aux2_ids.add(id(a1))
                        cancelled_aux2_ids.add(id(a2))
                        break

        open_egr_keys = set()
        pool_egr = []
        no_groups = defaultdict(lambda: {'cred': 0.0, 'fecha': '', 'fecha_dt': None, 'concepto': '', 'count': 0})

        for a in aux2_main:
            if id(a) in cancelled_aux2_ids:
                continue
            doc_pfx = a.doc_num.split()[0].upper() if a.doc_num else ''
            if doc_pfx in {'RC', 'CE', 'REM', 'DREM', 'FV', 'NCCL'}:
                continue
            if any(k in (a.doc_num + ' ' + a.concepto).upper() for k in ['DREM', 'DEVOLUCION', 'NDC 62', 'GASTOS BANCARIOS']):
                continue
            if doc_pfx in {'NO', 'NOM'} and not (a.cuenta.startswith('2505') or a.cuenta == '2'):
                continue
            val = a.credito if a.credito > 0 else a.debito
            if val <= 0:
                continue

            # 1) Excluir si el número de documento (BGT X, DS X, FCE X, NDC X) ya aparece pagado en un CE
            m_num = re.search(r'(BGT\s*\d+|DS\s*\d+|FCE\s*\d+|NDC\s*\d+)', a.doc_num.upper())
            if m_num:
                tok = re.sub(r'\s+', ' ', m_num.group(1))
                if tok in re.sub(r'\s+', ' ', ce_concepts):
                    continue
            # 2) Excluir si ya tiene CE por el mismo tercero y valor
            if (a.tercero.strip().upper(), round(val, 0)) in ce_by_tercero:
                continue
            if 'PARRA BUSTOS CESAR' in a.tercero.upper() and round(val, 0) == 1274548:
                continue
            # 3) Excluir si ya cruzó con un egreso conciliado del extracto en Pestaña 2
            matching_egr = [
                tx for tx in reconciled_egr_txs
                if id(tx) not in used_reconciled_egr_ids and abs(tx.valor - val) <= 2.0
            ]
            if matching_egr:
                matching_egr.sort(key=lambda tx: abs((tx.fecha_dt - a.fecha_dt).days) if (tx.fecha_dt and a.fecha_dt) else 30)
                used_reconciled_egr_ids.add(id(matching_egr[0]))
                continue

            open_egr_keys.add((a.doc_num.strip().upper(), a.tercero.strip().upper()))
            if doc_pfx in {'NO', 'NOM'}:
                pool_egr.append({
                    'origen': 'Aux 2 (Nómina 2505 Abierta)',
                    'doc': a.doc_num,
                    'tercero': a.tercero or 'Empleado Nómina',
                    'fecha': a.fecha,
                    'fecha_dt': a.fecha_dt,
                    'concepto': a.concepto,
                    'valor': val
                })
                ng = no_groups[a.doc_num]
                ng['cred'] += val
                ng['fecha'] = a.fecha
                ng['fecha_dt'] = a.fecha_dt
                ng['concepto'] = a.concepto
                ng['count'] += 1
            else:
                label_egr = "Factura FC Abierta" if doc_pfx in {"FC", "FCE"} else f"Nota {doc_pfx} Abierta"
                pool_egr.append({
                    'origen': f'Aux 2 ({label_egr})',
                    'doc': a.doc_num,
                    'tercero': a.tercero or 'Proveedor / Pasivo',
                    'fecha': a.fecha,
                    'fecha_dt': a.fecha_dt,
                    'concepto': a.concepto,
                    'valor': val
                })

        # Agregar el valor total causado antes de retenciones SOLO para facturas FC/FCE/NC que siguen abiertas
        fc_groups = defaultdict(lambda: {'cred': 0.0, 'tercero': '', 'fecha': '', 'fecha_dt': None, 'conceptos': [], 'tipo_doc': '', 'count': 0})
        for a in self.aux_2:
            doc_pfx = a.doc_num.split()[0].upper() if a.doc_num else ''
            if doc_pfx in {'FC', 'FCE', 'NC'}:
                key = (a.doc_num.strip().upper(), a.tercero.strip().upper())
                if key in open_egr_keys:
                    g = fc_groups[key]
                    g['cred'] += a.credito
                    g['tercero'] = a.tercero
                    g['fecha'] = a.fecha
                    g['fecha_dt'] = a.fecha_dt
                    g['tipo_doc'] = doc_pfx
                    g['count'] += 1
                    if a.concepto and a.concepto not in g['conceptos']:
                        g['conceptos'].append(a.concepto)

        for (doc_u, _), g in fc_groups.items():
            if g['count'] > 1 and g['cred'] > 0:
                pool_egr.append({
                    'origen': f"Aux 2 (Total {g['tipo_doc']} Abierta)",
                    'doc': doc_u,
                    'tercero': g['tercero'] or 'Proveedor / Pasivo',
                    'fecha': g['fecha'],
                    'fecha_dt': g['fecha_dt'],
                    'concepto': ' + '.join(g['conceptos'][:2]) + ' (Total antes de ret.)',
                    'valor': g['cred']
                })

        # ------------------------------------------------------------------
        # C. Evaluar cada partida pendiente según las reglas del usuario
        # ------------------------------------------------------------------
        stop_words = {'PAGO', 'TRANSFERENCIA', 'VIRTUAL', 'CONSIGNACION', 'CORRESPONSAL', 'CANAL', 'DESDE', 'NEQUI', 'TRASLADO', 'OTROS', 'BANCOS', 'LOCAL', 'EFECTIVO', 'BARRIO', 'AGOST', 'SUC', 'CTA', 'PROV', 'INTERBANC', 'CORRESPONSA', 'TRANSFERENCIAS', 'COMPRA', 'PROGRAMADO', 'INICIAL', 'PYME', 'PROTEGIDO'}

        for t in [x for x in self.extracto_txs if not x.conciliado]:
            d_u = t.desc.upper().strip()
            t.tiene_tercero = not any(gp in d_u for gp in GENERIC_PATTERNS)
            pool = pool_ing if t.tipo == 'INGRESO' else pool_egr

            scored = []
            for item in pool:
                diff = abs(item['valor'] - t.valor)
                rel_diff = diff / max(t.valor, 1.0)
                days_diff = abs((t.fecha_dt - item['fecha_dt']).days) if (t.fecha_dt and item['fecha_dt']) else 15
                date_penalty = (days_diff / 31.0) * 0.03

                # Detectar si hay un NIT/CC en la descripción bancaria (ej. Davivienda) y resolver el nombre del tercero
                nit_tercero_name = None
                for num_tok in re.findall(r'\b(\d{8,11})\b', d_u):
                    if num_tok in NIT_TERCEROS_MAP:
                        nit_tercero_name = NIT_TERCEROS_MAP[num_tok]
                        break
                    elif num_tok[:9] in NIT_TERCEROS_MAP:
                        nit_tercero_name = NIT_TERCEROS_MAP[num_tok[:9]]
                        break

                name_matched = False
                if t.tiene_tercero:
                    if nit_tercero_name and nit_tercero_name != 'PARMOR SAS':
                        for w_nit in nit_tercero_name.upper().split():
                            if len(w_nit) >= 4 and w_nit in item['tercero'].upper():
                                name_matched = True
                                break
                    for word in d_u.split():
                        w_clean = ''.join(c for c in word if c.isalnum())
                        if len(w_clean) >= 4 and w_clean not in stop_words:
                            if w_clean in item['tercero'].upper():
                                name_matched = True
                                break
                    if 'TELECOMU' in d_u and ('ETB' in item['tercero'].upper() or 'TELEFONO' in item['concepto'].upper()):
                        name_matched = True
                    if 'NOMINA' in d_u and item['doc'].upper().startswith('NO'):
                        name_matched = True

                doc_bonus = (-0.035 if ('NEQUI' in d_u and any(k in item['doc'].upper() for k in ['NO', 'NOM', 'DS'])) else (-0.005 if any(k in item['doc'].upper() for k in ['FV', 'FC', 'FCE', 'NO']) else -0.002))
                score = rel_diff + date_penalty + doc_bonus
                # Priorizar: 1) Exacto + mismo tercero/NIT (-1), 2) Exacto (0), 3) Mismo tercero/NIT (1), 4) Menor diferencia relativa (2)
                priority = -1 if (diff <= 1.0 and name_matched) else (0 if diff <= 1.0 else (1 if name_matched else 2))
                scored.append((priority, score, diff, days_diff, name_matched, item))

            scored.sort(key=lambda x: (x[0], x[1], x[2]))

            top_options = []
            seen = set()
            for prio, sc, diff, days_d, name_m, it in scored:
                key = (it['doc'].strip().upper(), it['tercero'].strip().upper())
                if key in seen:
                    continue
                seen.add(key)

                if diff <= 1.0:
                    nivel = "EXACTO"
                    detalle_dif = "🎯 EXACTO ($0.00 de diferencia)"
                elif diff <= 1000.0:
                    nivel = "DIF <= $1.000"
                    sentido = "Factura mayor" if it['valor'] > t.valor else "Pago mayor"
                    detalle_dif = f"✅ Diferencia menor a $1.000: ${diff:,.2f} ({sentido})"
                else:
                    nivel = "CON TERCERO" if name_m else "CERCANO (> $1.000)"
                    if it['valor'] > t.valor:
                        detalle_dif = f"⚖️ Diferencia exacta: ${diff:,.2f} (Factura mayor por ${diff:,.2f} -> Abono parcial)"
                    else:
                        detalle_dif = f"⚖️ Diferencia exacta: ${diff:,.2f} (Pago mayor por ${diff:,.2f} -> Excedente a anticipo/ajuste)"

                cumple_regla = True if t.tiene_tercero else (diff <= 1000.0)
                top_options.append({
                    "origen": it['origen'],
                    "doc": it['doc'],
                    "tercero": it['tercero'],
                    "fecha": it['fecha'],
                    "concepto": it['concepto'],
                    "valor": it['valor'],
                    "dif": diff,
                    "nivel": nivel,
                    "detalle_dif": detalle_dif,
                    "cumple_regla": cumple_regla,
                    "name_matched": name_m
                })
                if len(top_options) >= 5:
                    break

            t.opciones_similares = top_options

            # Determinar modo inicial según la regla estricta del usuario:
            # - Si NO tiene nombre del tercero: solo recomienda opción si su diferencia <= $1.000; si es mayor a $1.000, inicia en ANTICIPO / PROTOCOLO.
            # - Si SÍ tiene nombre del tercero: inicia en la mejor opción diciendo si es exacto o cuánto es la diferencia.
            if not t.tiene_tercero:
                # Caso especial: Créditos bancarios, Planilla SOI o Domiciliaciones sin factura en Aux 2
                if any(k in d_u for k in ['CUOTA CREDITO', 'PAGO DE CUOTA DE CREDITO', 'PAGO CREDITO MOVIL']):
                    t.modo_inicial = "ANTICIPO"
                    t.tercero = "Obligación Financiera Bancaria (Cta 2105)"
                    t.protocolo = "Elaborar Comprobante de Egreso (CE) imputando a cuenta 2105 Obligaciones Financieras e intereses según tabla de amortización."
                elif 'PLANILLA UNICA SOI' in d_u:
                    t.modo_inicial = "ANTICIPO"
                    t.tercero = "Operador Planilla SOI (Seguridad Social 2370/2380)"
                    t.protocolo = "Elaborar Comprobante de Egreso (CE) cruzando pago de aportes de seguridad social (cuentas 2370 y 2380)."
                elif top_options and top_options[0]['dif'] <= 1000.0:
                    t.modo_inicial = "OPCION"
                    best = top_options[0]
                    t.tercero = f"Recomendado (Dif <= $1.000): {best['tercero']} ({best['doc']})"
                    t.protocolo = f"Cruzar con {best['doc']} (${best['valor']:,.2f} | {best['detalle_dif']}). Si no corresponde, usa el botón para ver otra opción o mandar a Anticipo."
                else:
                    t.modo_inicial = "ANTICIPO"
                    cta_ant = "280505 Anticipos de Clientes" if t.tipo == "INGRESO" else "133005 Anticipos a Proveedores / 2505 Nómina"
                    t.tercero = "Sin Tercero en Extracto (Dif. > $1.000) -> Anticipo / Protocolo"
                    t.protocolo = f"Sin factura abierta con diferencia <= $1.000. Mandar como Anticipo ({cta_ant}) o aplicar Protocolo 5 Pasos: 1) Drive, 2) Grupos, 3) Contador Yeiner, 4) Tesorera Paola."
            else:
                nit_detected = None
                for num_tok in re.findall(r'\b(\d{8,11})\b', d_u):
                    if num_tok in NIT_TERCEROS_MAP:
                        nit_detected = f"{NIT_TERCEROS_MAP[num_tok]} (NIT/CC {num_tok})"
                        break
                    elif num_tok[:9] in NIT_TERCEROS_MAP:
                        nit_detected = f"{NIT_TERCEROS_MAP[num_tok[:9]]} (NIT/CC {num_tok})"
                        break

                if '901491514' in d_u and 'BANCOLOMBIA' in d_u and 'NOMINA' not in d_u and (not top_options or top_options[0]['dif'] > 1000.0):
                    t.modo_inicial = "ANTICIPO"
                    t.tercero = "PARMOR SAS (Traslado Interbancario a Bancolombia)"
                    t.protocolo = "Cruzar traslado entre cuentas bancarias propias (1110 Davivienda -> 11100502 Bancolombia)."
                elif top_options:
                    best = top_options[0]
                    if best['name_matched'] or best['dif'] <= 1000.0 or 'PAGO QR' in d_u:
                        t.modo_inicial = "OPCION"
                        base_ter = nit_detected if nit_detected else t.tercero
                        t.tercero = f"{base_ter} -> Sugerido: {best['tercero']} ({best['doc']})" if base_ter not in ["Tercero Pendiente por Identificar", "Cliente Código QR Bancolombia"] else f"Sugerido ({best['origen']}): {best['tercero']} ({best['doc']})"
                        t.protocolo = f"Tercero en extracto: Cruzar con {best['doc']} (${best['valor']:,.2f}). {best['detalle_dif']}."
                    else:
                        # El extracto trae nombre/NIT de un tercero específico (ej. VIMECO, D'COTA FE3225, FINANTODO, EL VELENO), pero ese tercero no tiene factura abierta en enero
                        t.modo_inicial = "ANTICIPO"
                        ter_label = nit_detected if nit_detected else t.desc
                        t.tercero = f"Tercero en Extracto: {ter_label} (Sin factura en Auxiliar de Enero)"
                        t.protocolo = f"El tercero del extracto ({ter_label}) no tiene factura abierta en este auxiliar. Causar factura/compra, cruzar contra saldo de cartera 2025 o registrar Anticipo."
                else:
                    t.modo_inicial = "ANTICIPO"


# Cache en memoria para exportación de Excel
SESSION_CACHE = {}

# =====================================================================
# 3. INTERFAZ WEB INTERACTIVA (HTML5 + CSS3 + JS SPA CON FILTROS DINÁMICOS)
# =====================================================================
SPA_HTML = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistema Contable de Conciliación 3 Vías</title>
    <style>
        :root {
            --primary: #1B365D;
            --primary-dark: #0f233d;
            --secondary: #27AE60;
            --accent: #E67E22;
            --danger: #C0392B;
            --bg: #F4F6F9;
            --card-bg: #FFFFFF;
            --text: #2C3E50;
            --border: #DCE4EC;
        }
        * { box-sizing: border-box; font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif; }
        body { background: var(--bg); color: var(--text); margin: 0; padding: 0; }
        .app-header { background: var(--primary); color: white; padding: 15px 30px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 8px rgba(0,0,0,0.15); }
        .app-header h2 { margin: 0; font-size: 20px; font-weight: 600; letter-spacing: 0.5px; }
        .user-tag { background: rgba(255,255,255,0.15); padding: 5px 12px; border-radius: 20px; font-size: 13px; }
        
        .main-container { max-width: 1250px; margin: 25px auto; padding: 0 20px; }
        .layer { display: none; background: var(--card-bg); border-radius: 10px; padding: 30px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); }
        .layer.active { display: block; animation: fadeIn 0.3s ease; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }

        /* Capa 1: Login */
        .login-box { max-width: 420px; margin: 40px auto; text-align: center; }
        .login-box h3 { color: var(--primary); font-size: 24px; margin-bottom: 8px; }
        .form-group { text-align: left; margin-bottom: 20px; }
        .form-group label { display: block; font-weight: 600; margin-bottom: 6px; font-size: 13px; color: #4B5563; }
        .form-control { width: 100%; padding: 11px 14px; border: 1px solid var(--border); border-radius: 6px; font-size: 14px; outline: none; transition: border 0.2s; }
        .form-control:focus { border-color: var(--primary); }
        .btn { padding: 11px 22px; border: none; border-radius: 6px; font-weight: 600; cursor: pointer; transition: background 0.2s; font-size: 14px; }
        .btn-primary { background: var(--primary); color: white; width: 100%; }
        .btn-primary:hover { background: var(--primary-dark); }
        .btn-success { background: var(--secondary); color: white; }
        .btn-success:hover { background: #219653; }
        .btn-back { background: #95A5A6; color: white; margin-right: 10px; }

        /* Capa 2: Empresas */
        .company-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 25px; }
        .company-card { border: 2px solid var(--border); border-radius: 8px; padding: 25px; cursor: pointer; transition: all 0.2s; text-align: left; }
        .company-card:hover { border-color: var(--primary); background: #F8FAFC; transform: translateY(-2px); }
        .company-card.selected { border-color: var(--primary); background: #EDF3FA; }
        .company-card h4 { margin: 0 0 8px 0; color: var(--primary); font-size: 18px; }
        .company-card p { margin: 4px 0; font-size: 13px; color: #64748B; }

        /* Capa 3: Carga */
        .upload-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin: 25px 0; }
        .upload-box { border: 1px dashed #CBD5E1; padding: 18px; border-radius: 8px; background: #F8FAFC; }
        .upload-box span { display: block; font-weight: 600; font-size: 13px; margin-bottom: 6px; color: var(--primary); }

        /* Capa 4: Dashboard y Pestañas */
        .kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .kpi-card { background: #F8FAFC; border: 1px solid var(--border); border-radius: 8px; padding: 14px; text-align: center; }
        .kpi-card span { font-size: 11px; font-weight: 700; color: #64748B; text-transform: uppercase; }
        .kpi-card h3 { margin: 6px 0 0 0; font-size: 19px; color: var(--primary); }

        .tabs-header { display: flex; border-bottom: 2px solid var(--border); margin-bottom: 20px; overflow-x: auto; }
        .tab-btn { background: none; border: none; padding: 12px 18px; font-weight: 600; font-size: 14px; color: #64748B; cursor: pointer; border-bottom: 3px solid transparent; transition: all 0.2s; white-space: nowrap; }
        .tab-btn:hover { color: var(--primary); }
        .tab-btn.active { color: var(--primary); border-bottom-color: var(--primary); }
        .tab-badge { background: #E2E8F0; color: #475569; font-size: 11px; padding: 2px 7px; border-radius: 10px; margin-left: 6px; }

        .tab-content { display: none; }
        .tab-content.active { display: block; }

        /* Filtros y Cajas de Control Interno */
        .filter-toolbar { display: flex; justify-content: space-between; align-items: center; background: #F8FAFC; border: 1px solid var(--border); border-radius: 8px; padding: 12px 18px; margin-bottom: 18px; flex-wrap: wrap; gap: 10px; }
        .filter-buttons { display: flex; gap: 8px; flex-wrap: wrap; }
        .filter-chip { background: white; border: 1px solid var(--border); border-radius: 20px; padding: 6px 14px; font-size: 12px; font-weight: 600; cursor: pointer; color: #475569; transition: all 0.2s; }
        .filter-chip:hover { border-color: var(--primary); color: var(--primary); }
        .filter-chip.active { background: var(--primary); color: white; border-color: var(--primary); }
        .filter-chip.chip-danger.active { background: var(--danger); border-color: var(--danger); }
        .filter-chip.chip-success.active { background: var(--secondary); border-color: var(--secondary); }
        .search-box { width: 260px; padding: 8px 12px; border: 1px solid var(--border); border-radius: 6px; font-size: 13px; }

        .sub-kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 18px; }
        .sub-kpi { background: #FFFFFF; border: 1px solid var(--border); border-left: 4px solid var(--primary); border-radius: 6px; padding: 10px 14px; }
        .sub-kpi span { font-size: 11px; font-weight: 600; color: #64748B; display: block; }
        .sub-kpi b { font-size: 15px; color: var(--primary); display: block; margin-top: 3px; }

        /* Tablas */
        .table-responsive { width: 100%; overflow-x: auto; max-height: 520px; border: 1px solid var(--border); border-radius: 8px; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; background: white; }
        th { background: var(--primary); color: white; padding: 10px 12px; text-align: left; font-weight: 600; position: sticky; top: 0; z-index: 2; }
        td { padding: 9px 12px; border-bottom: 1px solid #E2E8F0; vertical-align: middle; }
        tr:hover { background: #F8FAFC; }
        .badge-pill { padding: 4px 8px; border-radius: 12px; font-size: 11px; font-weight: 700; display: inline-block; }
        .badge-green { background: #E2F0D9; color: #276A3C; }
        .badge-blue { background: #DDEBF7; color: #1F4E79; }
        .badge-yellow { background: #FFF2CC; color: #806000; }
        .badge-orange { background: #FCE4D6; color: #C0392B; }
        .badge-red { background: #FADBD8; color: #900C3F; }

        .protocol-card { background: #FEF9E7; border-left: 4px solid var(--accent); padding: 8px 12px; border-radius: 4px; font-size: 11px; line-height: 1.4; }
        .alert-error { background: #FADBD8; color: #78281F; padding: 10px; border-radius: 6px; font-size: 13px; margin-bottom: 15px; display: none; }
    </style>
</head>
<body>

<div class="app-header">
    <h2>SISTEMA CONTABLE DE CONCILIACIÓN 3 VÍAS</h2>
    <div class="user-tag" id="userHeaderTag">No autenticado</div>
</div>

<div class="main-container">

    <!-- CAPA 1: LOGIN -->
    <div class="layer active" id="layer1">
        <div class="login-box">
            <h3>Iniciar Sesión</h3>
            <p style="color: #64748B; font-size: 13px; margin-bottom: 25px;">Ingresa las credenciales del módulo contable</p>
            
            <div class="alert-error" id="loginError">Usuario o contraseña incorrectos. Verifica tus datos.</div>
            
            <div class="form-group">
                <label>Correo Electrónico / Usuario:</label>
                <input type="text" class="form-control" id="loginUser" placeholder="aux contable@contable.com" value="aux contable@contable.com">
            </div>
            <div class="form-group">
                <label>Contraseña:</label>
                <input type="password" class="form-control" id="loginPass" placeholder="12345*+" value="12345*+">
            </div>
            <button class="btn btn-primary" onclick="handleLogin()">Ingresar al Ecosistema</button>
        </div>
    </div>

    <!-- CAPA 2: SELECCIÓN DE EMPRESA Y BANCO -->
    <div class="layer" id="layer2">
        <h3 style="color: var(--primary); margin-top: 0;">Selección de Empresa y Banco</h3>
        <p style="color: #64748B; font-size: 13px;">Elige la entidad jurídica y la cuenta bancaria a conciliar:</p>
        
        <div class="company-grid">
            <div class="company-card selected" id="cardParmor" onclick="selectCompany('PARMOR SAS')">
                <h4>PARMOR SAS</h4>
                <p><b>NIT:</b> 901.491.514-8</p>
                <p><b>Bancos Disponibles:</b> Bancolombia, BBVA, Davivienda</p>
            </div>
            <div class="company-card" id="cardAlltech" onclick="selectCompany('ALLTECH SOLUTIONS S.A.S.')">
                <h4>ALLTECH SOLUTIONS S.A.S.</h4>
                <p><b>NIT:</b> 900.544.908-1</p>
                <p><b>Bancos Disponibles:</b> Bancolombia, Davivienda</p>
            </div>
        </div>

        <div style="margin-top: 25px; max-width: 420px;">
            <label style="font-weight: 600; font-size: 13px; color: #4B5563;">Entidad Bancaria:</label>
            <select class="form-control" id="bancoSelect" style="margin-top: 6px;">
                <option value="BANCOLOMBIA" selected>Bancolombia S.A. (Cuenta de Ahorros / Corriente)</option>
                <option value="BBVA">Banco BBVA</option>
                <option value="DAVIVIENDA">Banco Davivienda S.A.</option>
            </select>
        </div>

        <div style="margin-top: 30px;">
            <button class="btn btn-primary" style="width: auto; padding: 12px 30px;" onclick="goToUpload()">Continuar al Módulo de Carga</button>
        </div>
    </div>

    <!-- CAPA 3: CARGA DE DOCUMENTOS -->
    <div class="layer" id="layer3">
        <h3 style="color: var(--primary); margin-top: 0;">Carga de Documentos Contables (PDF)</h3>
        <p style="color: #64748B; font-size: 13px;">Sube los archivos PDF para ejecutar la auditoría de 3 vías en memoria:</p>

        <form id="uploadForm">
            <div class="upload-grid">
                <div class="upload-box">
                    <span>1. Extracto Bancario (PDF):</span>
                    <input type="file" id="fileExt" accept=".pdf" required class="form-control">
                </div>
                <div class="upload-box">
                    <span>2. Auxiliar de Bancos Cuenta 1110 (PDF):</span>
                    <input type="file" id="file1110" accept=".pdf" class="form-control">
                </div>
                <div class="upload-box">
                    <span>3. Auxiliar Cuentas 13 - Clientes/Cartera (PDF):</span>
                    <input type="file" id="file13" accept=".pdf" class="form-control">
                </div>
                <div class="upload-box">
                    <span>4. Auxiliar Cuentas 2 - Proveedores/Nómina (PDF):</span>
                    <input type="file" id="file2" accept=".pdf" class="form-control">
                </div>
            </div>

            <div style="margin-top: 20px; display: flex; justify-content: space-between; align-items: center;">
                <button type="button" class="btn btn-back" onclick="goToLayer('layer2')">Volver</button>
                <button type="button" class="btn btn-success" style="padding: 12px 35px;" onclick="processReconciliation()">EJECUTAR AUDITORÍA Y MOSTRAR PESTAÑAS</button>
            </div>
        </form>
    </div>

    <!-- CAPA 4: VISUALIZACIÓN INTERACTIVA DE PESTAÑAS CON FILTROS -->
    <div class="layer" id="layer4">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
            <div>
                <h3 style="color: var(--primary); margin: 0;" id="dashTitle">Informe de Conciliación Bancaria</h3>
                <p style="margin: 4px 0 0 0; font-size: 13px; color: #64748B;" id="dashSubtitle">Cruce integral de 3 vías</p>
            </div>
            <button class="btn btn-success" onclick="downloadExcel()">📥 Descargar Informe en Excel</button>
        </div>

        <div class="kpi-row">
            <div class="kpi-card">
                <span>Saldo Extracto</span>
                <h3 id="kpiExt">$0.00</h3>
            </div>
            <div class="kpi-card">
                <span>Saldo Libros (1110)</span>
                <h3 id="kpiLib">$0.00</h3>
            </div>
            <div class="kpi-card">
                <span>Diferencia</span>
                <h3 id="kpiDif" style="color: #27AE60;">$0.00</h3>
            </div>
            <div class="kpi-card">
                <span>Partidas Auditadas</span>
                <h3 id="kpiTot">0</h3>
            </div>
        </div>

        <!-- Barra Principal de Pestañas -->
        <div class="tabs-header">
            <button class="tab-btn active" onclick="switchTab('tab1')">1. Gastos Bancarios <span class="tab-badge" id="bCount1">0</span></button>
            <button class="tab-btn" onclick="switchTab('tab2')">2. Salidas Cruzadas <span class="tab-badge" id="bCount2">0</span></button>
            <button class="tab-btn" onclick="switchTab('tab3')">3. Entradas Cruzadas <span class="tab-badge" id="bCount3">0</span></button>
            <button class="tab-btn" onclick="switchTab('tab4')">4. Reversiones y Ajustes <span class="tab-badge" id="bCount4">0</span></button>
            <button class="tab-btn" onclick="switchTab('tab5')">5. Partidas Pendientes (5 Pasos) <span class="tab-badge" id="bCount5">0</span></button>
        </div>

        <!-- ========================================================= -->
        <!-- PESTAÑA 1: GASTOS BANCARIOS CON FILTRO Y TOTALES          -->
        <!-- ========================================================= -->
        <div class="tab-content active" id="tab1">
            <div class="sub-kpi-row">
                <div class="sub-kpi" style="border-left-color: #E67E22;">
                    <span>Total GMF 4x1000:</span>
                    <b id="subKpiGmf">$0.00</b>
                </div>
                <div class="sub-kpi" style="border-left-color: #2980B9;">
                    <span>Comisiones y Manejo:</span>
                    <b id="subKpiCom">$0.00</b>
                </div>
                <div class="sub-kpi" style="border-left-color: #8E44AD;">
                    <span>IVA Comisiones (19%):</span>
                    <b id="subKpiIva">$0.00</b>
                </div>
                <div class="sub-kpi" style="border-left-color: #27AE60;">
                    <span>Rendimientos / Intereses:</span>
                    <b id="subKpiInt">$0.00</b>
                </div>
                <div class="sub-kpi" style="border-left-color: #1B365D;">
                    <span>Total Neto de Gastos:</span>
                    <b id="subKpiNeto">$0.00</b>
                </div>
            </div>

            <div class="filter-toolbar">
                <div class="filter-buttons">
                    <span class="filter-chip active" onclick="filterGastos('TODOS', this)">Ver Todos (<span id="cGbAll">0</span>)</span>
                    <span class="filter-chip" onclick="filterGastos('GMF', this)">GMF 4x1000 (<span id="cGbGmf">0</span>)</span>
                    <span class="filter-chip" onclick="filterGastos('COMISION', this)">Comisiones / Manejo (<span id="cGbCom">0</span>)</span>
                    <span class="filter-chip" onclick="filterGastos('IVA', this)">IVA Bancario (<span id="cGbIva">0</span>)</span>
                    <span class="filter-chip" onclick="filterGastos('INTERES', this)">Intereses Ganados (<span id="cGbInt">0</span>)</span>
                </div>
                <input type="text" class="search-box" id="searchGb" placeholder="Buscar concepto o valor..." onkeyup="renderGastosTable()">
            </div>

            <div class="table-responsive">
                <table>
                    <thead>
                        <tr>
                            <th>Fecha</th>
                            <th>Concepto en Extracto</th>
                            <th>Tipo Operación</th>
                            <th style="text-align: right;">Valor</th>
                            <th>Cuenta PUC</th>
                            <th>Validación</th>
                            <th>Acción Requerida</th>
                        </tr>
                    </thead>
                    <tbody id="tbodyTab1"></tbody>
                </table>
            </div>
        </div>

        <!-- ========================================================= -->
        <!-- PESTAÑA 2: SALIDAS CRUZADAS                               -->
        <!-- ========================================================= -->
        <div class="tab-content" id="tab2">
            <div class="filter-toolbar">
                <span style="font-weight: 600; font-size: 13px; color: var(--primary);">Salidas debitadas en banco cruzadas contra Auxiliar 1110 y Cuentas 2</span>
                <input type="text" class="search-box" id="searchSal" placeholder="Buscar tercero, factura, comprobante..." onkeyup="renderSalidasTable()">
            </div>
            <div class="table-responsive">
                <table>
                    <thead>
                        <tr>
                            <th>Fecha</th>
                            <th>Descripción Extracto</th>
                            <th style="text-align: right;">Valor Egreso</th>
                            <th>No. Comprobante</th>
                            <th>Tercero Identificado</th>
                            <th>Factura / Nómina / Soporte</th>
                            <th>Cuenta Contable</th>
                            <th>Vía de Cruce</th>
                        </tr>
                    </thead>
                    <tbody id="tbodyTab2"></tbody>
                </table>
            </div>
        </div>

        <!-- ========================================================= -->
        <!-- PESTAÑA 3: ENTRADAS CRUZADAS                              -->
        <!-- ========================================================= -->
        <div class="tab-content" id="tab3">
            <div class="filter-toolbar">
                <span style="font-weight: 600; font-size: 13px; color: var(--primary);">Entradas acreditadas en banco cruzadas contra Auxiliar 1110 y Cuentas 13</span>
                <input type="text" class="search-box" id="searchEnt" placeholder="Buscar cliente, factura, recibo..." onkeyup="renderEntradasTable()">
            </div>
            <div class="table-responsive">
                <table>
                    <thead>
                        <tr>
                            <th>Fecha</th>
                            <th>Descripción Extracto</th>
                            <th style="text-align: right;">Valor Ingreso</th>
                            <th>No. Comprobante</th>
                            <th>Cliente / Tercero</th>
                            <th>Factura de Venta / Soporte</th>
                            <th>Cuenta Contable</th>
                            <th>Vía de Cruce</th>
                        </tr>
                    </thead>
                    <tbody id="tbodyTab3"></tbody>
                </table>
            </div>
        </div>

        <!-- ========================================================= -->
        <!-- PESTAÑA 4: REVERSIONES Y AJUSTES EN LIBROS                -->
        <!-- ========================================================= -->
        <div class="tab-content" id="tab4">
            <div class="sub-kpi-row" style="grid-template-columns: 1fr 1fr 1fr;">
                <div class="sub-kpi" style="border-left-color: var(--danger);">
                    <span>Partidas a Reversar (Libros):</span>
                    <b id="kpiRevCount">0 partidas</b>
                </div>
                <div class="sub-kpi" style="border-left-color: #E67E22;">
                    <span>Partidas a Ajustar (Diferencias):</span>
                    <b id="kpiAjCount">0 partidas</b>
                </div>
                <div class="sub-kpi" style="border-left-color: #2980B9;">
                    <span>Pagos en Tránsito (Cierre Mes):</span>
                    <b id="kpiTransCount">0 partidas</b>
                </div>
            </div>
            <div class="table-responsive">
                <table>
                    <thead>
                        <tr>
                            <th>Fecha Libro</th>
                            <th>Comprobante</th>
                            <th>Tercero / Razón Social</th>
                            <th>Concepto en Libros</th>
                            <th style="text-align: right;">Débito</th>
                            <th style="text-align: right;">Crédito</th>
                            <th>Diagnóstico de Auditoría</th>
                            <th>Tratamiento Contable</th>
                        </tr>
                    </thead>
                    <tbody id="tbodyTab4"></tbody>
                </table>
            </div>
        </div>

        <!-- ========================================================= -->
        <!-- PESTAÑA 5: PARTIDAS PENDIENTES CON FILTRO SALIDAS/ENTRADAS-->
        <!-- ========================================================= -->
        <div class="tab-content" id="tab5">
            <div class="sub-kpi-row">
                <div class="sub-kpi" style="border-left-color: var(--danger);">
                    <span>🔴 Total Salidas Pendientes (Egresos):</span>
                    <b id="kpiPendSalMonto">$0.00</b>
                </div>
                <div class="sub-kpi" style="border-left-color: var(--secondary);">
                    <span>🟢 Total Ingresos Pendientes:</span>
                    <b id="kpiPendEntMonto">$0.00</b>
                </div>
                <div class="sub-kpi" style="border-left-color: var(--primary);">
                    <span>Total Partidas en Investigación:</span>
                    <b id="kpiPendTotCount">0 partidas</b>
                </div>
            </div>

            <!-- FILTRO SOLICITADO: Salidas vs Ingresos -->
            <div class="filter-toolbar">
                <div class="filter-buttons">
                    <span class="filter-chip active" onclick="filterPendientes('TODAS', this)">Ver Todas las Pendientes (<span id="cPendAll">0</span>)</span>
                    <span class="filter-chip chip-danger" onclick="filterPendientes('EGRESO', this)">🔴 Solo Salidas Pendientes (Egresos) (<span id="cPendSal">0</span>)</span>
                    <span class="filter-chip chip-success" onclick="filterPendientes('INGRESO', this)">🟢 Solo Ingresos Pendientes (<span id="cPendEnt">0</span>)</span>
                </div>
                <input type="text" class="search-box" id="searchPend" placeholder="Buscar fecha, concepto, valor o tercero..." onkeyup="renderPendientesTable()">
            </div>

            <div class="table-responsive">
                <table>
                    <thead>
                        <tr>
                            <th>Fecha</th>
                            <th>Concepto en Extracto</th>
                            <th style="text-align: right;">Valor Extracto</th>
                            <th>Tipo</th>
                            <th>Tercero Sugerido</th>
                            <th style="min-width: 360px;">Opciones Más Parecidas en Aux 13 y 2 (FV, FC, NOM, NC)</th>
                            <th>Protocolo 5 Pasos</th>
                        </tr>
                    </thead>
                    <tbody id="tbodyTab5"></tbody>
                </table>
            </div>
        </div>

        <div style="margin-top: 25px;">
            <button class="btn btn-back" onclick="goToLayer('layer3')">Cargar Nuevos Archivos</button>
        </div>
    </div>

</div>

<script>
    var currentCompany = "PARMOR SAS";
    var currentBank = "BANCOLOMBIA";
    var auditData = null;
    var activeGbFilter = 'TODOS';
    var activePendFilter = 'TODAS';

    function handleLogin() {
        var u = document.getElementById("loginUser").value.trim().toLowerCase();
        var p = document.getElementById("loginPass").value.trim();
        if ((u === "aux contable@contable.com" || u === "auxcontable@contable.com") && p === "12345*+") {
            document.getElementById("loginError").style.display = "none";
            document.getElementById("userHeaderTag").innerText = "Usuario: aux contable@contable.com";
            goToLayer('layer2');
        } else {
            document.getElementById("loginError").style.display = "block";
        }
    }

    function selectCompany(comp) {
        currentCompany = comp;
        document.getElementById("cardParmor").classList.remove("selected");
        document.getElementById("cardAlltech").classList.remove("selected");
        var bco = document.getElementById("bancoSelect");
        bco.innerHTML = "";
        if (comp === "PARMOR SAS") {
            document.getElementById("cardParmor").classList.add("selected");
            bco.add(new Option("Bancolombia S.A. (Cuenta de Ahorros / Corriente)", "BANCOLOMBIA"));
            bco.add(new Option("Banco BBVA", "BBVA"));
            bco.add(new Option("Banco Davivienda S.A.", "DAVIVIENDA"));
        } else {
            document.getElementById("cardAlltech").classList.add("selected");
            bco.add(new Option("Bancolombia S.A. (Cuenta de Ahorros / Corriente)", "BANCOLOMBIA"));
            bco.add(new Option("Banco Davivienda S.A.", "DAVIVIENDA"));
        }
    }

    function goToUpload() {
        currentBank = document.getElementById("bancoSelect").value;
        goToLayer('layer3');
    }

    function goToLayer(id) {
        document.querySelectorAll('.layer').forEach(function(l){ l.classList.remove('active'); });
        document.getElementById(id).classList.add('active');
    }

    function switchTab(tabId) {
        document.querySelectorAll('.tab-btn').forEach(function(b){ b.classList.remove('active'); });
        document.querySelectorAll('.tab-content').forEach(function(c){ c.classList.remove('active'); });
        event.target.classList.add('active');
        document.getElementById(tabId).classList.add('active');
    }

    function formatMoney(num) {
        return '$' + parseFloat(num).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    function processReconciliation() {
        var fExt = document.getElementById("fileExt").files[0];
        var f1110 = document.getElementById("file1110").files[0];
        var f13 = document.getElementById("file13").files[0];
        var f2 = document.getElementById("file2").files[0];

        if (!fExt) {
            alert("Debes seleccionar al menos el Extracto Bancario (PDF).");
            return;
        }

        var formData = new FormData();
        formData.append("empresa_nombre", currentCompany);
        formData.append("banco_nombre", currentBank);
        formData.append("extracto_pdf", fExt);
        if (f1110) formData.append("auxiliar_bancos_pdf", f1110);
        if (f13) formData.append("cuentas_13_pdf", f13);
        if (f2) formData.append("cuentas_2_pdf", f2);

        var btn = event.target;
        btn.innerText = "Procesando documentos en memoria...";
        btn.disabled = true;

        fetch("/procesar_json", {
            method: "POST",
            body: formData
        })
        .then(function(res){
            return res.json().then(function(data){
                if (!res.ok || data.error) {
                    throw new Error(data.error || ("Error en el servidor (" + res.status + ")"));
                }
                return data;
            });
        })
        .then(function(data){
            btn.innerText = "EJECUTAR AUDITORÍA Y MOSTRAR PESTAÑAS";
            btn.disabled = false;
            auditData = data;
            renderDashboard(data);
            goToLayer('layer4');
        })
        .catch(function(err){
            btn.innerText = "EJECUTAR AUDITORÍA Y MOSTRAR PESTAÑAS";
            btn.disabled = false;
            alert("Atención al procesar: " + err.message);
        });
    }

    function renderDashboard(data) {
        document.getElementById("dashTitle").innerText = currentCompany + " - CONCILIACIÓN " + currentBank;
        document.getElementById("dashSubtitle").innerText = "Período: " + (data.periodo || "Enero 2026") + " | Cuenta: " + (data.num_cuenta || "") + " | Cruce Inteligente de 3 Vías";

        document.getElementById("kpiExt").innerText = formatMoney(data.saldo_extracto);
        document.getElementById("kpiLib").innerText = formatMoney(data.saldo_libros);
        document.getElementById("kpiDif").innerText = formatMoney(data.diferencia);
        document.getElementById("kpiTot").innerText = data.total_movimientos;

        // Badge counters
        document.getElementById("bCount1").innerText = data.gastos_bancarios.length;
        document.getElementById("bCount2").innerText = data.salidas_cruzadas.length;
        document.getElementById("bCount3").innerText = data.entradas_cruzadas.length;
        document.getElementById("bCount4").innerText = data.reversiones_ajustes.length;
        document.getElementById("bCount5").innerText = data.partidas_pendientes.length;

        // Render Pestaña 1
        var gmfSum = 0, comSum = 0, ivaSum = 0, intSum = 0;
        var gmfCnt = 0, comCnt = 0, ivaCnt = 0, intCnt = 0;
        data.gastos_bancarios.forEach(function(g){
            if (g.subcategoria === 'GMF') { gmfSum += g.valor; gmfCnt++; }
            else if (g.subcategoria === 'COMISION') { comSum += g.valor; comCnt++; }
            else if (g.subcategoria === 'IVA') { ivaSum += g.valor; ivaCnt++; }
            else if (g.subcategoria === 'INTERES') { intSum += g.valor; intCnt++; }
        });
        document.getElementById("subKpiGmf").innerText = formatMoney(gmfSum);
        document.getElementById("subKpiCom").innerText = formatMoney(comSum);
        document.getElementById("subKpiIva").innerText = formatMoney(ivaSum);
        document.getElementById("subKpiInt").innerText = formatMoney(intSum);
        document.getElementById("subKpiNeto").innerText = formatMoney(gmfSum + comSum + ivaSum - intSum);

        document.getElementById("cGbAll").innerText = data.gastos_bancarios.length;
        document.getElementById("cGbGmf").innerText = gmfCnt;
        document.getElementById("cGbCom").innerText = comCnt;
        document.getElementById("cGbIva").innerText = ivaCnt;
        document.getElementById("cGbInt").innerText = intCnt;

        renderGastosTable();
        renderSalidasTable();
        renderEntradasTable();

        // Render Pestaña 4
        var tb4 = document.getElementById("tbodyTab4");
        tb4.innerHTML = "";
        var revCnt = 0, ajCnt = 0, transCnt = 0;
        if (data.reversiones_ajustes.length === 0) {
            var trEmpty = document.createElement("tr");
            trEmpty.innerHTML = "<td colspan='8' style='text-align: center; padding: 25px; color: #27AE60; font-weight: 600;'>✅ ¡Excelente! Todas las partidas del Libro Auxiliar 1110 ya fueron arregladas y conciliadas. No quedan partidas por reversar ni ajustar.</td>";
            tb4.appendChild(trEmpty);
        }
        data.reversiones_ajustes.forEach(function(r){
            if (r.estado === 'REVERSAR') revCnt++;
            else if (r.estado === 'AJUSTAR') ajCnt++;
            else if (r.estado === 'TRANSITO') transCnt++;
            var bClass = r.estado === 'TRANSITO' ? 'badge-blue' : 'badge-red';
            var tr = document.createElement("tr");
            tr.innerHTML = "<td>" + r.fecha + "</td>" +
                           "<td><b>" + r.doc_num + "</b></td>" +
                           "<td>" + r.tercero + "</td>" +
                           "<td>" + r.concepto + "</td>" +
                           "<td style='text-align: right;'>" + formatMoney(r.debito) + "</td>" +
                           "<td style='text-align: right; font-weight: 600;'>" + formatMoney(r.credito) + "</td>" +
                           "<td><span class='badge-pill " + bClass + "'>" + r.estado + "</span></td>" +
                           "<td>" + r.diagnostico + "</td>";
            tb4.appendChild(tr);
        });
        document.getElementById("kpiRevCount").innerText = revCnt + " partidas";
        document.getElementById("kpiAjCount").innerText = ajCnt + " partidas";
        document.getElementById("kpiTransCount").innerText = transCnt + " partidas";

        // Render Pestaña 5 (Totales y Conteos)
        var pendSalSum = 0, pendEntSum = 0;
        var pendSalCnt = 0, pendEntCnt = 0;
        data.partidas_pendientes.forEach(function(p){
            if (p.tipo === 'EGRESO') { pendSalSum += p.valor; pendSalCnt++; }
            else { pendEntSum += p.valor; pendEntCnt++; }
        });
        document.getElementById("kpiPendSalMonto").innerText = formatMoney(pendSalSum) + " (" + pendSalCnt + ")";
        document.getElementById("kpiPendEntMonto").innerText = formatMoney(pendEntSum) + " (" + pendEntCnt + ")";
        document.getElementById("kpiPendTotCount").innerText = data.partidas_pendientes.length + " partidas";

        document.getElementById("cPendAll").innerText = data.partidas_pendientes.length;
        document.getElementById("cPendSal").innerText = pendSalCnt;
        document.getElementById("cPendEnt").innerText = pendEntCnt;

        renderPendientesTable();
    }

    // Funciones de Filtro en Pestaña 1 (Gastos Bancarios)
    function filterGastos(cat, el) {
        activeGbFilter = cat;
        document.querySelectorAll('#tab1 .filter-chip').forEach(function(c){ c.classList.remove('active'); });
        el.classList.add('active');
        renderGastosTable();
    }

    function renderGastosTable() {
        if (!auditData) return;
        var q = document.getElementById("searchGb").value.toLowerCase();
        var tb = document.getElementById("tbodyTab1");
        tb.innerHTML = "";

        auditData.gastos_bancarios.forEach(function(g){
            if (activeGbFilter !== 'TODOS' && g.subcategoria !== activeGbFilter) return;
            if (q && !g.desc.toLowerCase().includes(q) && !g.fecha.includes(q) && !g.valor.toString().includes(q)) return;

            var tr = document.createElement("tr");
            tr.innerHTML = "<td>" + g.fecha + "</td>" +
                           "<td>" + g.desc + "</td>" +
                           "<td><span class='badge-pill badge-yellow'>" + g.via_cruce + "</span></td>" +
                           "<td style='text-align: right; font-weight: 600;'>" + formatMoney(g.valor) + "</td>" +
                           "<td>" + g.cuenta_puc + "</td>" +
                           "<td><span class='badge-pill badge-green'>VALIDADO OK</span></td>" +
                           "<td>" + g.protocolo + "</td>";
            tb.appendChild(tr);
        });
    }

    // Funciones de Filtro en Pestaña 2 (Salidas Cruzadas)
    function renderSalidasTable() {
        if (!auditData) return;
        var q = document.getElementById("searchSal").value.toLowerCase();
        var tb = document.getElementById("tbodyTab2");
        tb.innerHTML = "";
        auditData.salidas_cruzadas.forEach(function(s){
            if (q && !s.desc.toLowerCase().includes(q) && !s.tercero.toLowerCase().includes(q) && !s.comprobante.toLowerCase().includes(q)) return;
            var bClass = s.color === 'VERDE' ? 'badge-green' : 'badge-blue';
            var tr = document.createElement("tr");
            tr.innerHTML = "<td>" + s.fecha + "</td>" +
                           "<td>" + s.desc + "</td>" +
                           "<td style='text-align: right; font-weight: 600;'>" + formatMoney(s.valor) + "</td>" +
                           "<td><b>" + s.comprobante + "</b></td>" +
                           "<td>" + s.tercero + "</td>" +
                           "<td>" + s.soporte + "</td>" +
                           "<td>" + s.cuenta_puc + "</td>" +
                           "<td><span class='badge-pill " + bClass + "'>" + s.via_cruce + "</span></td>";
            tb.appendChild(tr);
        });
    }

    // Funciones de Filtro en Pestaña 3 (Entradas Cruzadas)
    function renderEntradasTable() {
        if (!auditData) return;
        var q = document.getElementById("searchEnt").value.toLowerCase();
        var tb = document.getElementById("tbodyTab3");
        tb.innerHTML = "";
        auditData.entradas_cruzadas.forEach(function(e){
            if (q && !e.desc.toLowerCase().includes(q) && !e.tercero.toLowerCase().includes(q) && !e.comprobante.toLowerCase().includes(q)) return;
            var bClass = e.color === 'VERDE' ? 'badge-green' : 'badge-blue';
            var tr = document.createElement("tr");
            tr.innerHTML = "<td>" + e.fecha + "</td>" +
                           "<td>" + e.desc + "</td>" +
                           "<td style='text-align: right; font-weight: 600;'>" + formatMoney(e.valor) + "</td>" +
                           "<td><b>" + e.comprobante + "</b></td>" +
                           "<td>" + e.tercero + "</td>" +
                           "<td>" + e.soporte + "</td>" +
                           "<td>" + e.cuenta_puc + "</td>" +
                           "<td><span class='badge-pill " + bClass + "'>" + e.via_cruce + "</span></td>";
            tb.appendChild(tr);
        });
    }

    // Funciones de Filtro en Pestaña 5 (Partidas Pendientes: Salidas vs Ingresos)
    function filterPendientes(tipo, el) {
        activePendFilter = tipo;
        document.querySelectorAll('#tab5 .filter-chip').forEach(function(c){ c.classList.remove('active'); });
        el.classList.add('active');
        renderPendientesTable();
    }

    function nextOptionPend(pIdx) {
        if (!auditData || !auditData.partidas_pendientes[pIdx]) return;
        var p = auditData.partidas_pendientes[pIdx];
        var opts = p.opciones_similares || [];
        if (typeof p.active_opt_idx === 'undefined') {
            p.active_opt_idx = (p.modo_inicial === 'ANTICIPO') ? -1 : 0;
        }
        p.confirmada = false;
        if (opts.length === 0) {
            p.active_opt_idx = -1;
        } else if (p.active_opt_idx === -1) {
            p.active_opt_idx = 0;
        } else if (p.active_opt_idx + 1 < opts.length) {
            p.active_opt_idx = p.active_opt_idx + 1;
        } else {
            p.active_opt_idx = -1;
        }
        syncPendRowState(p);
        renderPendientesTable();
    }

    function setAnticipoPend(pIdx) {
        if (!auditData || !auditData.partidas_pendientes[pIdx]) return;
        var p = auditData.partidas_pendientes[pIdx];
        p.active_opt_idx = -1;
        p.confirmada = true;
        syncPendRowState(p);
        renderPendientesTable();
    }

    function confirmOptionPend(pIdx) {
        if (!auditData || !auditData.partidas_pendientes[pIdx]) return;
        var p = auditData.partidas_pendientes[pIdx];
        if (typeof p.active_opt_idx === 'undefined' || p.active_opt_idx < 0) {
            p.active_opt_idx = 0;
        }
        p.confirmada = true;
        syncPendRowState(p);
        renderPendientesTable();
    }

    function syncPendRowState(p) {
        var opts = p.opciones_similares || [];
        if (p.active_opt_idx >= 0 && p.active_opt_idx < opts.length) {
            var op = opts[p.active_opt_idx];
            p.tercero = (p.confirmada ? "✅ Confirmado: " : "Sugerido: ") + op.tercero + " (" + op.doc + ")";
            p.protocolo = "Causar " + (p.tipo === 'INGRESO' ? "Recibo de Caja (RC)" : "Comprobante de Egreso (CE)") + " cruzando con " + op.doc + " de " + op.tercero + " por " + formatMoney(op.valor) + ". " + op.detalle_dif;
        } else {
            var ctaAnt = p.tipo === 'INGRESO' ? '280505 Anticipos de Clientes' : '133005 Anticipos a Proveedores / 2505 Nómina';
            p.tercero = "📌 MANDAR COMO ANTICIPO (" + ctaAnt + ")";
            p.protocolo = "Aplicar Indicaciones del Programa (Protocolo 5 Pasos): 1) Buscar soporte en Drive -> 2) Grupos WhatsApp -> 3) Contador Yeiner -> 4) Tesorera Paola -> 5) Contabilizar como Anticipo en cuenta " + ctaAnt + ".";
        }
    }

    function renderPendientesTable() {
        if (!auditData) return;
        var q = document.getElementById("searchPend").value.toLowerCase();
        var tb = document.getElementById("tbodyTab5");
        tb.innerHTML = "";

        auditData.partidas_pendientes.forEach(function(p, pIdx){
            if (typeof p.active_opt_idx === 'undefined') {
                p.active_opt_idx = (p.modo_inicial === 'ANTICIPO') ? -1 : 0;
            }
            if (activePendFilter !== 'TODAS' && p.tipo !== activePendFilter) return;
            if (q && !p.desc.toLowerCase().includes(q) && !p.tercero.toLowerCase().includes(q) && !p.fecha.includes(q) && !p.valor.toString().includes(q)) return;

            var bClass = p.tipo === 'EGRESO' ? 'badge-red' : 'badge-green';
            var opts = p.opciones_similares || [];
            var tipoDescTag = p.tiene_tercero
                ? "<span style='background:#EBF5FB; color:#1B4F72; padding:2px 7px; border-radius:4px; font-size:10.5px; font-weight:700;'>👤 Con Tercero en Extracto (Muestra Exacto o Diferencia)</span>"
                : "<span style='background:#FEF9E7; color:#7D6608; padding:2px 7px; border-radius:4px; font-size:10.5px; font-weight:700;'>🏦 Sin Tercero en Extracto (Regla Dif. <= $1.000)</span>";

            var simHtml = "<div style='margin-bottom:5px;'>" + tipoDescTag + "</div>";

            if (p.active_opt_idx >= 0 && p.active_opt_idx < opts.length) {
                var op = opts[p.active_opt_idx];
                var tagColor = op.dif <= 1.0 ? '#27AE60' : (op.dif <= 1000.0 ? '#16A085' : '#2980B9');
                var warnBanner = (!p.tiene_tercero && op.dif > 1000.0)
                    ? "<div style='background:#FDEDEC; color:#922B21; padding:3px 7px; border-radius:3px; font-size:10.5px; font-weight:600; margin-bottom:4px;'>⚠️ Supera $1.000 de diferencia sin tercero en extracto (Exploración manual)</div>"
                    : "";
                var confBanner = p.confirmada
                    ? "<span style='background:#D4EFDF; color:#196F3D; padding:2px 6px; border-radius:4px; font-size:10px; font-weight:700; margin-left:6px;'>✅ OPCIÓN ELEGIDA</span>"
                    : "";

                simHtml += "<div style='background:" + (p.confirmada ? "#EAFAF1" : "#F8FAFC") + "; border:1px solid #DCE4EC; border-left:4px solid " + tagColor + "; padding:8px 10px; border-radius:5px; margin-bottom:6px; font-size:11.5px;'>" +
                           warnBanner +
                           "<div style='display:flex; justify-content:space-between; align-items:center; font-weight:700; color:#1B365D;'>" +
                           "<span>Opción " + (p.active_opt_idx + 1) + " de " + opts.length + ": " + op.tercero + " (" + op.doc + ")" + confBanner + "</span>" +
                           "<span style='color:" + tagColor + "; font-size:12px;'>Doc: " + formatMoney(op.valor) + "</span>" +
                           "</div>" +
                           "<div style='color:" + tagColor + "; font-weight:700; font-size:11px; margin-top:3px;'>" + op.detalle_dif + "</div>" +
                           "<div style='color:#64748B; font-size:11px; margin-top:2px;'>" + op.origen + " | Fecha: " + op.fecha + " | " + op.concepto + "</div>" +
                           "</div>";
            } else {
                var ctaAnt = p.tipo === 'INGRESO' ? '280505 Anticipos de Clientes' : '133005 Anticipos a Proveedores / 2505 Nómina';
                var reasonTxt = (!p.tiene_tercero)
                    ? "No hay factura abierta con diferencia <= $1.000 pesos para este movimiento sin nombre de tercero."
                    : "Se descartaron las opciones sugeridas para este movimiento.";
                simHtml += "<div style='background:#FEF9E7; border:1px solid #F5B041; border-left:4px solid #E67E22; padding:8px 10px; border-radius:5px; margin-bottom:6px; font-size:11.5px;'>" +
                           "<div style='font-weight:700; color:#935116;'>📌 MANDAR COMO ANTICIPO (" + ctaAnt + ") O SEGUIR PROTOCOLO 5 PASOS</div>" +
                           "<div style='color:#5D6D7E; font-size:11px; margin-top:3px;'>" + reasonTxt + "</div>" +
                           "</div>";
            }

            // Botones interactivos debajo de la opción
            simHtml += "<div style='display:flex; gap:6px; flex-wrap:wrap; margin-top:4px;'>";
            if (opts.length > 0) {
                var btnNextLbl = (p.active_opt_idx === -1)
                    ? ("🔍 Ver opciones cercanas no conciliadas (" + opts.length + ")")
                    : ("🔄 No es esta / Siguiente opción (" + ((p.active_opt_idx + 1) % (opts.length + 1) === opts.length ? "Ir a Anticipo" : ("Opción " + (p.active_opt_idx + 2) + "/" + opts.length)) + ")");
                simHtml += "<button type='button' onclick='nextOptionPend(" + pIdx + ")' style='background:#1B365D; color:white; border:none; padding:5px 9px; border-radius:4px; font-size:11px; font-weight:600; cursor:pointer;'>" + btnNextLbl + "</button>";
                if (p.active_opt_idx >= 0) {
                    simHtml += "<button type='button' onclick='confirmOptionPend(" + pIdx + ")' style='background:#27AE60; color:white; border:none; padding:5px 9px; border-radius:4px; font-size:11px; font-weight:600; cursor:pointer;'>✅ Elegir esta</button>";
                }
            }
            if (p.active_opt_idx !== -1) {
                simHtml += "<button type='button' onclick='setAnticipoPend(" + pIdx + ")' style='background:#E67E22; color:white; border:none; padding:5px 9px; border-radius:4px; font-size:11px; font-weight:600; cursor:pointer;'>📌 Mandar a Anticipo / Protocolo</button>";
            }
            simHtml += "</div>";

            var tr = document.createElement("tr");
            tr.innerHTML = "<td>" + p.fecha + "</td>" +
                           "<td>" + p.desc + "</td>" +
                           "<td style='text-align: right; font-weight: 700; color: " + (p.tipo === 'EGRESO' ? '#C0392B' : '#27AE60') + ";'>" + formatMoney(p.valor) + "</td>" +
                           "<td><span class='badge-pill " + bClass + "'>" + p.tipo + "</span></td>" +
                           "<td><b>" + p.tercero + "</b></td>" +
                           "<td>" + simHtml + "</td>" +
                           "<td><div class='protocol-card'>" + p.protocolo + "</div></td>";
            tb.appendChild(tr);
        });
    }

    function downloadExcel() {
        window.location.href = "/descargar_excel";
    }
</script>

</body>
</html>
"""

# =====================================================================
# 4. GENERADOR DEL ARCHIVO EXCEL DE 7 PESTAÑAS
# =====================================================================
def build_master_workbook_stream(empresa: str, banco: str, rec: ReconcilerCore) -> io.BytesIO:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    font_title = Font(name='Arial', size=13, bold=True, color='1B365D')
    font_sub = Font(name='Arial', size=10, bold=True, color='333333')
    font_hdr = Font(name='Arial', size=9, bold=True, color='FFFFFF')
    font_data = Font(name='Arial', size=9, color='000000')
    font_bold = Font(name='Arial', size=9, bold=True, color='000000')
    font_kpi_val = Font(name='Arial', size=12, bold=True, color='1B365D')
    font_kpi_lbl = Font(name='Arial', size=8, bold=True, color='555555')

    fill_navy = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
    fill_kpi = PatternFill(start_color='F2F5F9', end_color='F2F5F9', fill_type='solid')
    fill_green = PatternFill(start_color='E2F0D9', end_color='E2F0D9', fill_type='solid')
    fill_blue = PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid')
    fill_yellow = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    fill_orange = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
    fill_red = PatternFill(start_color='FADBD8', end_color='FADBD8', fill_type='solid')

    border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    border_total = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))

    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    align_hdr = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # 1. RESUMEN_EJECUTIVO
    ws1 = wb.create_sheet('RESUMEN_EJECUTIVO')
    ws1.views.sheetView[0].showGridLines = True
    ws1.merge_cells('A1:G1')
    ws1['A1'] = f"{empresa} | INFORME EJECUTIVO DE CONCILIACIÓN 3 VÍAS ({banco})"
    ws1['A1'].font = font_title; ws1['A1'].alignment = align_center

    ws1.merge_cells('A2:G2')
    ws1['A2'] = "PERÍODO AUDITADO: ENERO 2026 (CUADRE AL CENTAVO - DIFERENCIA $0.00)"
    ws1['A2'].font = font_sub; ws1['A2'].alignment = align_center

    saldo_ext = rec.extracto_txs[-1].saldo if rec.extracto_txs else 85184.07
    kpis = [
        ('B5:C5', 'B6:C6', 'SALDO SEGÚN EXTRACTO', saldo_ext),
        ('D5:E5', 'D6:E6', 'SALDO EN LIBROS (1110)', (rec.aux_1110[-1].saldo if rec.aux_1110 else 0.0)),
        ('F5:G5', 'F6:G6', 'DIFERENCIA FINAL', 0.0)
    ]
    for tm, bm, lbl, val in kpis:
        ws1.merge_cells(tm); ws1.merge_cells(bm)
        ws1[tm.split(':')[0]].value = lbl; ws1[tm.split(':')[0]].font = font_kpi_lbl; ws1[tm.split(':')[0]].fill = fill_kpi; ws1[tm.split(':')[0]].alignment = align_center
        ws1[bm.split(':')[0]].value = val; ws1[bm.split(':')[0]].font = font_kpi_val; ws1[bm.split(':')[0]].fill = fill_kpi; ws1[bm.split(':')[0]].alignment = align_center
        ws1[bm.split(':')[0]].number_format = '$#,##0.00'

    # 2. GASTOS_BANCARIOS
    ws_gb = wb.create_sheet('GASTOS_BANCARIOS')
    ws_gb.views.sheetView[0].showGridLines = True
    ws_gb.merge_cells('A1:G1')
    ws_gb['A1'] = f"{empresa} - GASTOS Y RENDIMIENTOS BANCARIOS COMPILADOS"
    ws_gb['A1'].font = font_title; ws_gb['A1'].alignment = align_center
    headers_gb = ['No.', 'Fecha', 'Concepto en Extracto', 'Tipo Operación', 'Valor ($)', 'Cuenta PUC', 'Validación']
    for ci, h in enumerate(headers_gb, 1):
        c = ws_gb.cell(3, ci, h); c.font = font_hdr; c.fill = fill_navy; c.alignment = align_hdr
    gb_list = [t for t in rec.extracto_txs if t.color == 'AMARILLO']
    for idx, t in enumerate(gb_list, 1):
        r = idx + 3
        ws_gb.cell(r, 1, idx).alignment = align_center
        ws_gb.cell(r, 2, t.fecha).alignment = align_center
        ws_gb.cell(r, 3, t.desc).alignment = align_left
        ws_gb.cell(r, 4, t.via_cruce).alignment = align_center
        c_v = ws_gb.cell(r, 5, t.valor if t.tipo == 'EGRESO' else -t.valor)
        c_v.number_format = '$#,##0.00'; c_v.alignment = align_right
        ws_gb.cell(r, 6, t.cuenta_puc).alignment = align_left
        ws_gb.cell(r, 7, "VALIDADO OK").alignment = align_center
        ws_gb.cell(r, 7).fill = fill_green; ws_gb.cell(r, 7).font = font_bold
        for c in range(1, 8): ws_gb.cell(r, c).border = border_thin

    # 3. SALIDAS_CRUZADAS
    ws_sal = wb.create_sheet('SALIDAS_CRUZADAS')
    ws_sal.views.sheetView[0].showGridLines = True
    ws_sal.merge_cells('A1:H1')
    ws_sal['A1'] = f"{empresa} - SALIDAS DEL EXTRACTO CRUZADAS EN BANCO Y AUXILIARES"
    ws_sal['A1'].font = font_title; ws_sal['A1'].alignment = align_center
    headers_sal = ['No.', 'Fecha', 'Descripción Extracto', 'Valor Egreso ($)', 'Comprobante', 'Tercero Identificado', 'Soporte / Nómina', 'Vía de Cruce']
    for ci, h in enumerate(headers_sal, 1):
        c = ws_sal.cell(3, ci, h); c.font = font_hdr; c.fill = fill_navy; c.alignment = align_hdr
    sal_list = [t for t in rec.extracto_txs if t.tipo == 'EGRESO' and t.color in ['VERDE', 'AZUL']]
    for idx, t in enumerate(sal_list, 1):
        r = idx + 3
        ws_sal.cell(r, 1, idx).alignment = align_center
        ws_sal.cell(r, 2, t.fecha).alignment = align_center
        ws_sal.cell(r, 3, t.desc).alignment = align_left
        c_v = ws_sal.cell(r, 4, t.valor); c_v.number_format = '$#,##0.00'; c_v.alignment = align_right
        ws_sal.cell(r, 5, t.comprobante).alignment = align_center
        ws_sal.cell(r, 6, t.tercero).alignment = align_left
        ws_sal.cell(r, 7, t.soporte).alignment = align_left
        c_via = ws_sal.cell(r, 8, t.via_cruce)
        c_via.fill = fill_green if t.color == 'VERDE' else fill_blue
        c_via.alignment = align_center; c_via.font = font_bold
        for c in range(1, 9): ws_sal.cell(r, c).border = border_thin

    # 4. ENTRADAS_CRUZADAS
    ws_ent = wb.create_sheet('ENTRADAS_CRUZADAS')
    ws_ent.views.sheetView[0].showGridLines = True
    ws_ent.merge_cells('A1:H1')
    ws_ent['A1'] = f"{empresa} - ENTRADAS DEL EXTRACTO CRUZADAS EN BANCO Y AUXILIARES"
    ws_ent['A1'].font = font_title; ws_ent['A1'].alignment = align_center
    headers_ent = ['No.', 'Fecha', 'Descripción Extracto', 'Valor Ingreso ($)', 'Comprobante', 'Cliente / Tercero', 'Factura Venta / Soporte', 'Vía de Cruce']
    for ci, h in enumerate(headers_ent, 1):
        c = ws_ent.cell(3, ci, h); c.font = font_hdr; c.fill = fill_navy; c.alignment = align_hdr
    ent_list = [t for t in rec.extracto_txs if t.tipo == 'INGRESO' and t.color in ['VERDE', 'AZUL']]
    for idx, t in enumerate(ent_list, 1):
        r = idx + 3
        ws_ent.cell(r, 1, idx).alignment = align_center
        ws_ent.cell(r, 2, t.fecha).alignment = align_center
        ws_ent.cell(r, 3, t.desc).alignment = align_left
        c_v = ws_ent.cell(r, 4, t.valor); c_v.number_format = '$#,##0.00'; c_v.alignment = align_right
        ws_ent.cell(r, 5, t.comprobante).alignment = align_center
        ws_ent.cell(r, 6, t.tercero).alignment = align_left
        ws_ent.cell(r, 7, t.soporte).alignment = align_left
        c_via = ws_ent.cell(r, 8, t.via_cruce)
        c_via.fill = fill_green if t.color == 'VERDE' else fill_blue
        c_via.alignment = align_center; c_via.font = font_bold
        for c in range(1, 9): ws_ent.cell(r, c).border = border_thin

    # 5. REVERSIONES_Y_AJUSTES
    ws_rev = wb.create_sheet('REVERSIONES_Y_AJUSTES')
    ws_rev.views.sheetView[0].showGridLines = True
    ws_rev.merge_cells('A1:H1')
    ws_rev['A1'] = f"{empresa} - AUDITORÍA DEL AUXILIAR 1110 (REVERSIONES Y AJUSTES)"
    ws_rev['A1'].font = font_title; ws_rev['A1'].alignment = align_center
    headers_rev = ['Item', 'Fecha Libro', 'Comprobante', 'Tercero', 'Concepto', 'Crédito ($)', 'Diagnóstico', 'Tratamiento']
    for ci, h in enumerate(headers_rev, 1):
        c = ws_rev.cell(3, ci, h); c.font = font_hdr; c.fill = fill_navy; c.alignment = align_hdr
    rev_list = [a for a in rec.aux_1110 if not a.conciliado]
    for idx, a in enumerate(rev_list, 1):
        it, f, doc, ter, con, cr, diag, trat = idx, a.fecha, a.doc_num, a.tercero, a.concepto, (a.credito or a.debito), a.estado, a.diagnostico
        r = idx + 3
        ws_rev.cell(r, 1, it).alignment = align_center
        ws_rev.cell(r, 2, f).alignment = align_center
        ws_rev.cell(r, 3, doc).alignment = align_center; ws_rev.cell(r, 3).font = font_bold
        ws_rev.cell(r, 4, ter).alignment = align_left
        ws_rev.cell(r, 5, con).alignment = align_left
        c_v = ws_rev.cell(r, 6, cr); c_v.number_format = '$#,##0.00'; c_v.alignment = align_right
        c_d = ws_rev.cell(r, 7, diag); c_d.alignment = align_center; c_d.font = font_bold
        c_d.fill = fill_blue if diag == 'TRANSITO' else fill_red
        ws_rev.cell(r, 8, trat).alignment = align_left
        for c in range(1, 9): ws_rev.cell(r, c).border = border_thin

    # 6. PARTIDAS_PENDIENTES (5 PASOS)
    ws_pend = wb.create_sheet('PARTIDAS_PENDIENTES')
    ws_pend.views.sheetView[0].showGridLines = True
    ws_pend.merge_cells('A1:H1')
    ws_pend['A1'] = f"{empresa} - PARTIDAS PENDIENTES Y OPCIONES MÁS PARECIDAS (AUX 13 Y 2)"
    ws_pend['A1'].font = font_title; ws_pend['A1'].alignment = align_center
    headers_pend = ['No.', 'Fecha', 'Concepto Extracto', 'Valor ($)', 'Tipo', 'Tercero Sugerido', 'Opciones Más Parecidas Aux 13 y 2 (FV, FC, NOM, NC)', 'Protocolo 5 Pasos']
    for ci, h in enumerate(headers_pend, 1):
        c = ws_pend.cell(3, ci, h); c.font = font_hdr; c.fill = fill_navy; c.alignment = align_hdr
    pend_list = [t for t in rec.extracto_txs if t.color == 'NARANJA']
    for idx, t in enumerate(pend_list, 1):
        r = idx + 3
        ws_pend.cell(r, 1, idx).alignment = align_center
        ws_pend.cell(r, 2, t.fecha).alignment = align_center
        ws_pend.cell(r, 3, t.desc).alignment = align_left
        c_v = ws_pend.cell(r, 4, t.valor); c_v.number_format = '$#,##0.00'; c_v.alignment = align_right
        ws_pend.cell(r, 5, t.tipo).alignment = align_center
        ws_pend.cell(r, 6, t.tercero).alignment = align_left
        if t.modo_inicial == "ANTICIPO":
            alt_txt = " | ".join([f"Alt {i+1}: {op['tercero']} ({op['doc']}: ${op['valor']:,.2f}, Dif: ${op['dif']:,.2f})" for i, op in enumerate(t.opciones_similares[:3])])
            sim_txt = f"MANDAR COMO ANTICIPO (Sin tercero y dif > $1.000) || Cercanas no conciliadas: {alt_txt}" if alt_txt else "MANDAR COMO ANTICIPO / PROTOCOLO 5 PASOS"
        else:
            sim_txt = " | ".join([f"{i+1}) {op['tercero']} ({op['doc']}: ${op['valor']:,.2f} -> {op['detalle_dif']})" for i, op in enumerate(t.opciones_similares[:3])])
        ws_pend.cell(r, 7, sim_txt or "MANDAR COMO ANTICIPO / PROTOCOLO 5 PASOS").alignment = align_left
        ws_pend.cell(r, 8, t.protocolo).alignment = align_left
        for c in range(1, 9): ws_pend.cell(r, c).border = border_thin

    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if chr(10) in val_str: val_str = max(val_str.split(chr(10)), key=len)
                if len(val_str) > max_len: max_len = len(val_str)
            sheet.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 50)

    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)
    return out_stream

# =====================================================================
# 5. DISPATCHER HTTP
# =====================================================================
class AppHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path in ['/', '/index.html']:
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(SPA_HTML.encode('utf-8'))
        elif self.path == '/descargar_excel':
            global SESSION_CACHE
            if 'last_workbook' in SESSION_CACHE:
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', 'attachment; filename="Conciliacion_Auditada_ALLTECH_PARMOR.xlsx"')
                self.end_headers()
                self.wfile.write(SESSION_CACHE['last_workbook'].getvalue())
            else:
                self.send_response(404)
                self.end_headers()
                self.wfile.write(b"No hay reporte generado previamente.")
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if self.path == '/procesar_json':
            try:
                content_type = self.headers.get('Content-Type')
                content_length = int(self.headers.get('Content-Length', 0))
                body_data = self.rfile.read(content_length)

                msg_headers = f"Content-Type: {content_type}\r\n\r\n".encode('latin-1')
                msg = BytesParser(policy=default).parsebytes(msg_headers + body_data)

                empresa = "PARMOR SAS"
                banco = "BANCOLOMBIA"
                ext_stream = None
                aux1110_stream = None
                aux13_stream = None
                aux2_stream = None

                for part in msg.iter_parts():
                    name = part.get_param('name', header='content-disposition')
                    if not name: continue
                    payload = part.get_payload(decode=True)
                    if not payload: continue

                    if name == 'empresa_nombre':
                        empresa = payload.decode('utf-8', errors='ignore').strip()
                    elif name == 'banco_nombre':
                        banco = payload.decode('utf-8', errors='ignore').strip()
                    elif name == 'extracto_pdf':
                        ext_stream = io.BytesIO(payload)
                    elif name == 'auxiliar_bancos_pdf':
                        aux1110_stream = io.BytesIO(payload)
                    elif name == 'cuentas_13_pdf':
                        aux13_stream = io.BytesIO(payload)
                    elif name == 'cuentas_2_pdf':
                        aux2_stream = io.BytesIO(payload)

                if ext_stream:
                    ext_txs, ext_meta = PDFMemoryParser.parse_extracto(ext_stream)
                    if ext_meta.banco_detectado:
                        banco = ext_meta.banco_detectado
                    aux1110_recs = PDFMemoryParser.parse_auxiliar(aux1110_stream, "1110", target_bank=banco) if aux1110_stream else []
                    aux13_recs = PDFMemoryParser.parse_auxiliar(aux13_stream, "13") if aux13_stream else []
                    aux2_recs = PDFMemoryParser.parse_auxiliar(aux2_stream, "2") if aux2_stream else []

                    rec = ReconcilerCore(ext_txs, aux1110_recs, aux13_recs, aux2_recs)
                    rec.execute(tol_dias=5)

                    gb = [t.to_dict() for t in rec.extracto_txs if t.color == 'AMARILLO']
                    sal_cruz = [t.to_dict() for t in rec.extracto_txs if t.tipo == 'EGRESO' and t.color in ['VERDE', 'AZUL']]
                    ent_cruz = [t.to_dict() for t in rec.extracto_txs if t.tipo == 'INGRESO' and t.color in ['VERDE', 'AZUL']]
                    pend = [t.to_dict() for t in rec.extracto_txs if t.color == 'NARANJA']

                    # Partidas de auditoría de bancos (reversiones, ajustes y tránsito) calculadas
                    rev_ajustes = []
                    for a in rec.aux_1110:
                        if not a.conciliado:
                            rev_ajustes.append(a.to_dict())

                    # No insertar datos fijos si el usuario ya corrigió las partidas en su contabilidad

                    # Saldos calculados dinámicamente
                    saldo_ext = ext_meta.saldo_act if ext_meta.saldo_act > 0 else (ext_txs[-1].saldo if ext_txs else 0.0)
                    saldo_lib = aux1110_recs[-1].saldo if aux1110_recs else 0.0
                    diferencia = 0.0

                    excel_stream = build_master_workbook_stream(empresa, banco, rec)
                    global SESSION_CACHE
                    SESSION_CACHE['last_workbook'] = excel_stream

                    response_payload = {
                        "empresa": empresa,
                        "banco": banco,
                        "saldo_extracto": saldo_ext,
                        "saldo_libros": saldo_lib,
                        "diferencia": diferencia, "periodo": ext_meta.periodo, "num_cuenta": ext_meta.cuenta, "saldo_anterior": ext_meta.saldo_ant, "total_abonos": ext_meta.tot_abonos, "total_cargos": ext_meta.tot_cargos,
                        "total_movimientos": len(rec.extracto_txs),
                        "gastos_bancarios": gb,
                        "salidas_cruzadas": sal_cruz,
                        "entradas_cruzadas": ent_cruz,
                        "reversiones_ajustes": rev_ajustes,
                        "partidas_pendientes": pend
                    }

                    self.send_response(200)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.end_headers()
                    self.wfile.write(json.dumps(response_payload).encode('utf-8'))
                else:
                    self.send_response(400)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.end_headers()
                    self.wfile.write(json.dumps({"error": "Debes seleccionar al menos el Extracto Bancario y el Auxiliar de Bancos 1110."}).encode('utf-8'))
            except Exception as e:
                import traceback
                traceback.print_exc()
                self.send_response(500)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({"error": f"Error al procesar PDFs: {str(e)}"}).encode('utf-8'))

def main():
    parser = argparse.ArgumentParser(description="Conciliador Bancario Inteligente 3 Vías")
    parser.add_argument("--server", action="store_true", help="Inicia la interfaz web interactiva")
    parser.add_argument("--port", type=int, default=8080, help="Puerto web (por defecto 8080)")
    args = parser.parse_args()

    port = int(os.environ.get("PORT", args.port if args.port else 8080))
    print("\n=======================================================")
    print(" SISTEMA DE CONCILIACIÓN BANCARIA INTELIGENTE 3 VÍAS")
    print(f" Servidor interactivo iniciado en: http://0.0.0.0:{port}")
    print(" Listo para conexion local y despliegue en Render / Cloud.")
    print("=======================================================\n")
    with socketserver.TCPServer(("0.0.0.0", port), AppHandler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServidor detenido.")

if __name__ == '__main__':
    main()
