import unittest

from services import ai_invoice_service as svc


FIXTURE_TEXT = """Industrial Farmacéutica Cantabria, S.A.
N.I.F.: A-39000914
Barrio Solía 30, La Concha de Villaescusa
(39690), Cantabria, España
IMPUESTOS
TOTAL BRUTO
BASE IMPONIBLE
%
I.V.A.
%
REC.EQUIV
TOTAL
1.171,34
(1)
21,00
245,98
1.171,34
1.417,32 EUR
FACTURA Nº
TIPO FAC
FECHA
CLIENTE Nº
PR. ZONA
Nº ALBARAN
N.I.F.
26009701
RI
03/02/2026
235044
VENCIMIENTO:
60 dias fecha factura
04/04/2026
"""

HENRY_SCHEIN_FIXTURE = """REF
CTD.
PEDIDA
CTD.
SERVIDA
DTO. %
DESCRIPCIÓN
PRECIO
UNITARIO
PRECIO
NETO
BASE
I.V.A.
TOTAL
Nº ALBARÁN
Henry Schein Medical SL Avda. De la Albufera 153 – 28038 Madrid  CIF: B-02679538
FECHA FACTURA
N° FACTURA
22-04-26
A136372
RECIBO 15 DIAS FECHA FACTURA
BASE EXENTA
TOTAL A PAGAR
BASE 2
BASE 1
TOTAL BASE
63.67
92.44
EURO
77.67
14.00
BASE 3 (4.0%)
TOTAL I.V.A
I.V.A 1
I.V.A 2
I.V.A 3
13.37
14.77
1.40
"""

HENRY_SCHEIN_FULL_DATE_FIXTURE = """FECHA PEDIDO
C.I.F. / N.I.F.
FECHA FACTURA
N° FACTURA
SU REFERENCIA
N° CLIENTE
F A C T U R A
674259
A136372
17936479
702882
27676
A2
22-04-26
22-04-26
BASE EXENTA
TOTAL A PAGAR
BASE 2
BASE 1
TOTAL BASE
63.67
92.44
EURO
77.67
14.00
BASE 3 (4.0%)
TOTAL I.V.A
I.V.A 1
I.V.A 2
I.V.A 3
13.37
14.77
1.40
"""

HENRY_SCHEIN_REPEATED_SUMMARY_FIXTURE = """BASE EXENTA
TOTAL A PAGAR
BASE 2
BASE 1
I.V.A 1
I.V.A 2
TOTAL BASE
TOTAL I.V.A
BASE 3 (4.0%)
I.V.A 3
INCOTERM
REF
OTRA PAGINA
BASE EXENTA
TOTAL A PAGAR
BASE 2
BASE 1
TOTAL BASE
190.29
257.24
EURO
214.83
24.54
BASE 3 (4.0%)
TOTAL I.V.A
I.V.A 1
I.V.A 2
I.V.A 3
39.96
42.41
2.45
"""

CANTABRIA_MULTIVAT_FIXTURE = """Industrial Farmacéutica Cantabria, S.A.
IMPUESTOS
TOTAL BRUTO
BASE IMPONIBLE
%
I.V.A.
%
REC.EQUIV
TOTAL
111,35
(1)
10,00
11,14

462,36
(2)
21,00
97,10

573,71
681,95 EUR
VENCIMIENTO:
60 dias fecha factura
27/06/2026
"""

CANTABRIA_SINGLE_DUE_WITH_FOOTER_DATE_FIXTURE = """Industrial Farmacéutica Cantabria, S.A.
FACTURA Nº
FECHA
26038145
08/05/2026
FORMA DE PAGO:
Giro Bancario a la cuenta
OBSERVACIONES:
BANCO:
**** **** **** ** ******1816
VENCIMIENTO:
60 dias fecha factura
07/07/2026
Cantabria Labs es una marca registrada de Industrial Farmacéutica Cantabria, S.A.
Inscrita en el Reg. Merc. de Santander, Fecha 12/11/42, Libro 22 de Sociedades, Folio 105, Hoja 1.151, Inscripción 1ª
"""

YSONUT_FIXTURE = """YSONUT SLU - Calle Maldonado 50. 28006, Madrid - NIF: ESB61741112
Factura Simplificada
Fecha entrega 30/04/2026
Cliente : 30036373
CIF/NIF: B05410667
Nº Factura: INP-26FIG1-011569
Fecha Factura: 30/04/2026
Total referencias
Base impuesto
Tasa
Importe impuesto
Tipo de IVA
 10,00
IVA Reducido
 169,47
 16,95
Total base imponible
 169,47
 16,95
Total IVA
Total factura
 186,42
EUR
 186,42
NETO A PAGAR
Forma de pago
Vencimiento
Importe
CLIENTE Domiciliación bancaria
 93,21
31/05/2026
CC:ES3900492626802114161816
CLIENTE Domiciliación bancaria
 93,21
15/06/2026
CC:ES3900492626802114161816
"""

YSONUT_LONG_FIXTURE = """YSONUT SLU - Calle Maldonado 50. 28006, Madrid - NIF: ESB61741112
Factura Simplificada
CERVANTES 25 BJ
46007  VALENCIA
KALOS HEALTH AND BEAUTY S.R.L.
Fecha entrega 30/04/2026
Cliente : 30036373
CIF/NIF: B05410667
Nº Factura: INP-26FIG1-011546
Fecha Factura: 30/04/2026
Nº Albaran :ECCFIG12618179
Pag:
1/2
Referencia
Descripción
Unidades
Precio Neto
Precio Total
% IVA
Precio
Dto1 - Dto2
C080NB06
 23,36
 10,00
 11,680
 2
 20,00
INOVANCE MAGNESIUM
20,00 - 27,00
C446NB06
 32,39
 10,00
 16,193
 2
 27,73
INOVANCE VITA K2-D3
20,00 - 27,00
C528NA06
 52,98
 10,00
 26,492
 2
 45,36
INOVANCE HYALUROVANCE COLLAGENE RADIANCE
20,00 - 27,00
34FLYC528A62601
 0,000
 10
 0,00
DOC- FLYER HYALUROVANCE RADIANCE 2026
0,00
C375NA06
 74,06
 10,00
 24,687
 3
 42,27
INOVANCE HYALUROVANCE
20,00 - 27,00
C448NB06
 52,98
 10,00
 26,492
 2
 45,36
INOVANCE DERMOVANCE M
20,00 - 27,00
C482LB06
 105,97
 10,00
 26,492
 4
 45,36
INOVANCE CAPIVANCE 180
20,00 - 27,00
KALOS HEALTH AND BEAUTY S.R.L.
AVDA GUILLEM DE CASTRO 8 PTA 20
46001  VALENCIA
Dirección entrega:
 341,74
Total referencias
Base impuesto
Tasa
Importe impuesto
Tipo de IVA
 10,00
IVA Reducido
 341,74
 34,17
Total base imponible
 341,74
 34,17
Total IVA
Total factura
 375,91
EUR
 375,91
NETO A PAGAR
Forma de pago
Vencimiento
Importe
CLIENTE Domiciliación bancaria
 187,96
31/05/2026
CC:ES3900492626802114161816
CLIENTE Domiciliación bancaria
 187,95
15/06/2026
CC:ES3900492626802114161816
C/Maldonado, 50 28006-MADRID         C/Provença, 302 bajos 08038-BARCELONA          C/Holanda,16 17600-FIGUERES
Tel. 915 90 39 40                                      Tel. 934 88 10 96                                                       Fax 972 512190
www.ysonut.com
YSONUT SLU NIF/EORI: ESB61741112
YSONUT SLU - Calle Maldonado 50. 28006, Madrid - NIF: ESB61741112
Factura Simplificada
CERVANTES 25 BJ
46007  VALENCIA
KALOS HEALTH AND BEAUTY S.R.L.
Fecha entrega 30/04/2026
Cliente : 30036373
CIF/NIF: B05410667
Nº Factura: INP-26FIG1-011546
Fecha Factura: 30/04/2026
LEY DE PROTECCIÓN DE DATOS
Los datos de carácter personal  facilitados serán tratados por YSONUT , S.L.U. con NIF B61741112 de acuerdo con lo dispuesto en el Reglamento (UE) 2016/679 del Parlamento
Europeo y del Consejo, de 27 de abril de 2016, relativo a la protección de las personas físicas en lo que respecta al tratamiento de datos personales y a la libre circulación de los
mismos.
Los datos facilitados serán tratados por el tiempo necesario para el cumplimiento de las finalidades objeto de tratamiento, mientras no se oponga al mismo y por el tiempo
necesario para el cumplimiento de las obligaciones legales del responsable.
Los datos no serán cedidos ni comunicados a terceros, salvo en los supuestos legalmente establecidos.
Le recordamos que tiene derecho a ejercer los derechos de acceso, rectificación, cancelación, limitación, oposición y portabilidad de manera gratuita mediante correo electrónico a:
rgpd@ysonut.com o bien en la siguiente dirección: C/ Maldonado, 50, 28006 - Madrid (Madrid) y de solicitar la tutela de la Agencia Española de Protección de datos en
www.aepd.es.
DERECHO DE DESISTIMIENTO
D/Dª_______________________________________, mayor de edad y con DNI num__________________________y domicilio a efecto de notificaciones en, mediante el presente
documento, y dentro del plazo legal de 14 días a contar desde la recepción de los productos, vengo a..
Consecuentemente, procedo a la devolución de los productos a YSONUT SLU , en su domicilio sito en CL MALDONADO 50 28006 MADRID , siendo los gastos de devolución de
mi cuenta.Asimismo, interesa que, una vez recibidos los productos procedan a la devolución de los importes abonados por los mismos.
C/Maldonado, 50 28006-MADRID         C/Provença, 302 bajos 08038-BARCELONA          C/Holanda,16 17600-FIGUERES
Tel. 915 90 39 40                                      Tel. 934 88 10 96                                                       Fax 972 512190
www.ysonut.com
YSONUT SLU NIF/EORI: ESB61741112
"""

RAIOLA_FIXTURE = """PAGADA
Raiola Networks, S.L.
Avda. de Magoi nº 66
Semisótano Dcha
27002 Lugo España
CIF: B27453489
Factura nºE202605-53115
Fecha de la Factura: 11/05/2026
Fecha de Vencimiento: 16/05/2026
Facturado a
KALOS HEALTH AND BEAUTY SL
Descripción
Total
Hosting Base SSD 2.0 - fullfacevalencia.es (16/05/2026 - 15/05/2027)
€98.95 EUR
Subtotal
€98.95 EUR
21.00% IVA
€20.78 EUR
Saldo
€0.00 EUR
Total
€119.73 EUR
Transacciones
Fecha de transacción
Forma de pago
ID Transacción
Total
11/05/2026
Tarjeta de Crédito o Débito
260511104624
€119.73 EUR
Pendiente
€0.00 EUR
PDF generado el 11/05/2026
"""

SKINTECH_FIXTURE = """30/04/2026
ESB05410667
C004609
100% A 30 DIAS
Basado en Pedidos de cliente 502600616. Basado en Entregas 302600753.
SEPA Direct Debit Clients
30/05/2026
 1.084,16
FACTURA
Mencionar en el pago
Cliente Nº
Fecha
Nº de factura
Dirección Envío
NIF
Condición de Pago
Vía de Pago
Vencimientos
Datos bancarios
Condiciones de entrega
Su referencia
Observaciones
Código
Cantidad
Precio/uni
Dto.
Importe
CALLE CERVANTES NUMERO 25
BAJO
46007  VALENCIA
ESPAÑA
Incoterm
Lote
Fecha vencim.
Descripción
Productos fabricados en España
KALOS HEALTH AND BEAUTY S.L.
CALLE CERVANTES NUMERO 25
BAJO
46007  VALENCIA
ESPAÑA
IB26400354
SKIN TECH PHARMA GROUP S.L.U. Inscrita en el Registro Mercantil de la Provincia de Girona Tomo 904 Folio 97, Hoja GI-16.953. NIF/TAX ID ES B 17.470.261
 302600753
30/04/2026
Subtotal
Dto.
Desglose
Base Imponible
% IVA
Importe IVA
Total Factura
Total Productos
Base Imponible
 896,00
 21,00%
 1.084,16
 188,16
 896,00
 1.084,16
 1.084,16
EUR
TOTAL IMPORTE PENDIENTE
Skin Tech Pharma Group, SLU
C/ Pla de l'estany,29
17486 Castelló d'Empúries - Girona (Spain)
Telf. 972455113
1
Página
En cumplimiento de lo establecido en el Reglamento (UE) 2016/679 del Parlamento Europeo y del Consejo, de 27 de abril de 2016, relativo a la protección de datos personales y la libre circulación de los mismos, le comunicamos que los datos que usted nos facilite quedarán incorporados y serán tratados en los
ficheros titularidad de SKIN TECH PHARMA GROUP, S.L. (NIF B17470261) con el fin de poderle prestar nuestros servicios y realizar la facturación de los mismos. Los datos proporcionados se conservarán durante el tiempo que dure la relación comercial o durante el tiempo que sea necesario para cumplir con
las obligaciones legales. Dichos datos no serán transferidos a terceros, a menos que exista una obligación legal de hacerlo. Usted puede ejercer sus derechos de acceso, rectificación, oposición, supresión, limitación del tratamiento, portabilidad y de no ser objeto de decisiones individualizadas ante SKIN TECH
PHARMA GROUP SL, Pla de l'Estany 29, 17486 - Castelló d'Empúries o en la dirección de correo electrónico gdpr@skintechpharmagroup.com, adjuntando copia de su DNI o documento equivalente.  Asimismo, puede presentar una reclamación ante la Agencia Española de Protección de Datos, C/ Jorge Juan, 6
– 28001 Madrid.
Impreso por SAP Business One
"""

VERISURE_FIXTURE = """C/ Priégola 2 - 28224, Pozuelo de Alarcón, Madrid – España. Tel.: 910 121 122. RNSP2737
C.I.F. A26106013 -
Factura Cuota Mayo 2026
76,85 €
Total a pagar
www.verisure.es
KALOS HEALTH AND BEAUTY SL
2605C00829192
01/05/2026
01/05/2026-31/05/2026
Cuenta Bancaria:
Factura Cuota Mayo 2026
Concepto
Cantidad
Importe unitario
Importe total
ALARMA ANTI-INTRUSION
1
4,00 €
4,00 €
ALARMA ZEROVISION ANTI-INTRUSION
1
45,51 €
45,51 €
DISPOSITIVOS ADICIONALES
1
13,00 €
13,00 €
SERVICIOS CONFORT
1
2,00 €
2,00 €
DESCUENTO EN CUOTA SERVICIO AMPLIACION
1
-3,00 €
-3,00 €
OFERTA CUOTA FIDELIZACION
1
-1,00 €
-1,00 €
60,51 €
Subtotal
12,71 €
73,22 €
Total factura
IVA  21 %
TOTAL A PAGAR
76,85 €
Pago aplazado servicio de instalación dispositivos ampliados (2 de 36). Factura: 26OT000117440
3,63 €
Inscrita en el Reg. Merc. de Madrid – Tomo 9454 – Libro de Sociedades 0 – Folio 75 – Hoja M – 151.950 – Inscripción 3º. A los efectos informativos oportunos, Securitas Direct España
S.A.U., como empresa inscrita en el Registro RII-AEE con el número 002846 y en el Registro RII-PYA con el número 655
"""

AUTONOMO_PHARMACY_FIXTURE = """Código
Descripción del Artículo
%Ap. P.V.P Bruto P.V.P Neto
Unid.
Importe
CRISTINA SIMÓ BESALDUCH
CORREOS, 14
VALENCIA (VALENCIA)
Cod.Farmacia: 38
N.I.F. 20903358S
Factura Número: Q000081/2026
KALOS HEALTH AND BEAUTY SL
C/CERVANTES Nº25 BAJO
VALENCIA
N.I.F: B05410667
Referencia: 1566
Nombre:
Dirección:
Población:
Fecha
Fecha: 13/05/26
%Iva
C.P.: 46007
E-Mail:
N/A
Operación nº B001207/2026
BASE IMPONIBLE:
TOTAL CUOTAS:
TOTAL FACTURA:
677,83 €
27,11 €
704,94 €
"""

MIXED_SUPPLIER_PRIORITY_FIXTURE = """JUAN PÉREZ GARCÍA
N.I.F. 12345678Z
KALOS HEALTH AND BEAUTY SL
N.I.F. B05410667
Proveedor:
SUMINISTROS CLÍNICOS LEVANTE, S.L.
Factura Número: A-25
"""

CONSENUR_FIXTURE = """Kalos Health And Beauty S.L
C.I.F.: B05410667
Vía CERVANTES 25 BAJO
46007 VALENCIA(Valencia)
España
Consenur Sanitarios, S.L.
C.I.F.: B86208824
Vía Calle Río Ebro, s/n Polígono Industrial Finanzauto
28500 Arganda del Rey, ARGANDA DEL REY (MADRID)
Tlfno: 900922903 - Fax:
Fecha:
07/05/2026
Nº Fact:
1004187E2676973
Vencimiento:
07/05/2026
Día de pago:
Concepto
DCS
Cant. Unidad
PrecioCant.
ImporteTipoImp.
Kalos Health And Beauty S.L - Vía CERVANTES 25 BAJO  46007 VALENCIA ( Valencia )
30/06/2026
Albaran:
Nº Pedido:
Cuota Trimestral Abril 2026 - Junio 2026
1,00
92,92 €
92,92 €
10,00%
Tratamiento Residuos Biosanitarios
TOTAL SERVICIOS:
92,92 €
Base Imponible: (10,00 %)
92,92 €
 Total IVA:
9,29 €
 Total Base Imponible:
92,92 €
 Total IVA:
9,29 €
 TOTAL FACTURA:
102,21 €
Forma de pago Pago domiciliado
Nº Cuenta: ES3900492626802114161816
Observaciones:
CONSENUR SANITARIOS, S.L.U. Inscrita en el Registro Mercantil de Madrid al Tomo 28.907,  Folio 93, Hoja M-520536, Inscripción 62.
1/1
"""

MEDIDERMA_CLIENT_TYPO_FIXTURE = """FACTURA
Cliente - Destinatario factura:
KARLOS HEALTH AND BEAUTY S.L.
C/ CERVANTES 25 BAJO,
46007 VALENCIA
Valencia
España
Cliente - Datos fiscales:
KARLOS HEALTH AND BEAUTY S.L.
C/ CERVANTES 25 BAJO,
46007 VALENCIA
Valencia
España
Nº cliente 138272 N.I.F. B05410667
Nº de pedido del cliente WS464687
Dirección Envío C/ CERVANTES 25 BAJO  VALENCIA 46007
Factura Nº  122602793 de fecha 13.05.2026
Albarán nº:  0081551593
Artículo
Material/Descripción
TOTAL EUR
Total EUR
133,10
MEDIDERMA, S.L.U. · N.I.F. ESB96188164 · Domicilio fiscal: C/ GRABADOR ESTEVE 3-1-1, 46004  VALENCIA (España)
"""

RENT_INVOICE_IRPF_COLUMNS_FIXTURE = """FACTURA 350
Alquiler mes actual finca urbana del local
BASE IMPONIBLE       % I.V.A.       I.V.A.
300,00               21,00          63,00
BASE I.R.P.F.        % I.R.P.F.     I.R.P.F.
300,00               19,00          57,00
TOTAL FACTURA: 306,00
"""

VAT_EXEMPT_CONFIRMING_FIXTURE = """FACTURA
LIQUID. COM./INT. DE FINANCIACION 16,95 EUR
Base Sujeta y Exenta 16,95 EUR
Total Importe 16,95 EUR
"""

PROFESSIONAL_INVOICE_WITHHOLDING_FIXTURE = """FACTURA N° NOMINA AGOSTO CHELO
DATOS DEL EMISOR
Mª Consuelo Sebastián Pastor · DNI: 25414619X
DATOS DEL CLIENTE
KALOS HEALTH AND BEAUTY SL · ID: B05410667
Fecha: 31 de agosto de 2026
Actividad médica Agosto.
Base imponible 2264,15
IVA (21%) 475,47
IRPF (15%) -339,62
TOTAL A PERCIBIR 2400,00
Importe sujeto a retención del 15% de IRPF.
"""


class TestAiInvoiceService(unittest.TestCase):
    def test_parse_eu_amount(self):
        cases = {
            "1.171,34": 1171.34,
            "1 171,34": 1171.34,
            "1417,32 EUR": 1417.32,
            "245,98": 245.98,
            "0,00": 0.0,
        }
        for raw, expected in cases.items():
            parsed = svc.parse_eu_amount(raw)
            self.assertIsNotNone(parsed)
            self.assertAlmostEqual(parsed, expected, places=2)

    def test_extract_tax_summary_from_text(self):
        summary = svc._extract_tax_summary_from_text(FIXTURE_TEXT)
        self.assertTrue(summary.get("found"))
        self.assertAlmostEqual(summary.get("base_amount"), 1171.34, places=2)
        self.assertAlmostEqual(summary.get("vat_rate"), 21.0, places=2)
        self.assertAlmostEqual(summary.get("vat_amount"), 245.98, places=2)
        self.assertAlmostEqual(summary.get("total_amount"), 1417.32, places=2)

    def test_tax_summary_overrides_llm(self):
        summary = svc._extract_tax_summary_from_text(FIXTURE_TEXT)
        base, vat, total, rate, source = svc._apply_tax_summary_override(
            FIXTURE_TEXT,
            971.34,
            245.98,
            1217.32,
            21.0,
            summary,
        )
        self.assertEqual(source, "regex_tax_summary")
        self.assertAlmostEqual(base, 1171.34, places=2)
        self.assertAlmostEqual(vat, 245.98, places=2)
        self.assertAlmostEqual(total, 1417.32, places=2)
        self.assertAlmostEqual(rate, 21.0, places=2)

        corrected = svc._reconcile_vat_breakdown(
            [{"rate": 21.0, "base": 971.36, "vat_amount": 245.98, "total": 1217.32}],
            base,
            vat,
            total,
            rate,
            source,
        )
        self.assertEqual(len(corrected), 1)
        self.assertAlmostEqual(corrected[0]["base"], 1171.34, places=2)
        self.assertAlmostEqual(corrected[0]["vat_amount"], 245.98, places=2)
        self.assertAlmostEqual(corrected[0]["total"], 1417.32, places=2)

    def test_invoice_and_payment_dates(self):
        invoice_date = svc._extract_invoice_date_from_text(FIXTURE_TEXT)
        self.assertEqual(invoice_date, "2026-02-03")
        payment_dates = svc._find_payment_dates_by_keywords(FIXTURE_TEXT, invoice_date)
        self.assertIn("2026-04-04", payment_dates)

    def test_confidence_score_mapping(self):
        self.assertAlmostEqual(svc._confidence_score_for_source("regex_tax_summary"), 0.98, places=2)
        self.assertAlmostEqual(svc._confidence_score_for_source("llm"), 0.85, places=2)
        self.assertAlmostEqual(svc._confidence_score_for_source("fallback"), 0.60, places=2)

    def test_llm_amounts_trustworthy(self):
        self.assertTrue(svc._is_llm_amounts_trustworthy(362.58, 21.0, 76.14, 438.72))
        self.assertFalse(svc._is_llm_amounts_trustworthy(21.0, 21.0, 21.0, 61.98))

    def test_normalize_multivat_rates_by_math(self):
        extracted = {
            "vat_breakdown": [
                {"base": 18.86, "vat_amount": 3.96, "rate": 4},
                {"base": 8.40, "vat_amount": 0.84, "rate": 21},
            ],
            "totals": {"base": None, "vat": None, "total": None},
        }
        normalized = svc.normalize_and_validate_amounts(extracted)
        breakdown = normalized["vat_breakdown"]
        self.assertEqual(len(breakdown), 2)
        rates = sorted([line["rate"] for line in breakdown])
        self.assertEqual(rates, [10.0, 21.0])
        self.assertIsNone(normalized["vat_rate"])
        self.assertAlmostEqual(normalized["base_amount"], 27.26, places=2)
        self.assertAlmostEqual(normalized["vat_amount"], 4.8, places=2)
        self.assertAlmostEqual(normalized["total_amount"], 32.06, places=2)

    def test_totals_reconciled_from_breakdown(self):
        extracted = {
            "totals": {"base": 100.0, "vat": 10.0, "total": 110.0},
            "vat_breakdown": [
                {"base": 50.0, "vat_amount": 10.5},
                {"base": 60.0, "vat_amount": 12.6},
            ],
        }
        normalized = svc.normalize_and_validate_amounts(extracted)
        self.assertAlmostEqual(normalized["base_amount"], 110.0, places=2)
        self.assertAlmostEqual(normalized["vat_amount"], 23.1, places=2)
        self.assertAlmostEqual(normalized["total_amount"], 133.1, places=2)

    def test_partial_when_totals_incoherent(self):
        extracted = {"totals": {"base": 100.0, "vat": 10.0, "total": 50.0}}
        normalized = svc.normalize_and_validate_amounts(extracted)
        self.assertEqual(normalized["analysis_status"], "partial")
        self.assertIsNone(normalized["base_amount"])
        self.assertIsNone(normalized["vat_amount"])
        self.assertIsNone(normalized["total_amount"])

    def test_payment_terms_15_days(self):
        text = "RECIBO 15 DIAS FECHA FACTURA"
        terms = svc.extract_payment_terms_days(text)
        self.assertEqual(terms, 15)
        invoice_date = "2020-02-26"
        payment_date = (svc.date.fromisoformat(invoice_date) + svc.timedelta(days=terms)).isoformat()
        self.assertEqual(payment_date, "2020-03-12")

    def test_zero_vat_totals_are_preserved(self):
        extracted = {
            "supplier": "Google Cloud EMEA Limited",
            "invoice_date": "2026-01-31",
            "totals": {"base": 32.40, "vat": 0.00, "total": 32.40},
            "vat_breakdown": [
                {"base": 32.40, "vat_amount": 0.00, "rate": None},
            ],
        }
        normalized = svc.normalize_and_validate_amounts(extracted)
        self.assertEqual(normalized["analysis_status"], "ok")
        self.assertAlmostEqual(normalized["base_amount"], 32.40, places=2)
        self.assertAlmostEqual(normalized["vat_amount"], 0.00, places=2)
        self.assertAlmostEqual(normalized["total_amount"], 32.40, places=2)
        self.assertAlmostEqual(normalized["vat_rate"], 0.0, places=2)
        self.assertEqual(len(normalized["vat_breakdown"]), 1)

    def test_extract_numbered_tax_summary_from_text(self):
        summary = svc._extract_tax_summary_from_text(HENRY_SCHEIN_FIXTURE)
        self.assertTrue(summary.get("found"))
        self.assertAlmostEqual(summary.get("base_amount"), 77.67, places=2)
        self.assertAlmostEqual(summary.get("vat_amount"), 14.77, places=2)
        self.assertAlmostEqual(summary.get("total_amount"), 92.44, places=2)
        self.assertEqual(len(summary.get("breakdown") or []), 2)

    def test_extract_repeated_numbered_tax_summary_prefers_complete_block(self):
        summary = svc._extract_tax_summary_from_text(HENRY_SCHEIN_REPEATED_SUMMARY_FIXTURE)
        self.assertTrue(summary.get("found"))
        self.assertAlmostEqual(summary.get("base_amount"), 214.83, places=2)
        self.assertAlmostEqual(summary.get("vat_amount"), 42.41, places=2)
        self.assertAlmostEqual(summary.get("total_amount"), 257.24, places=2)

    def test_extract_invoice_date_dd_mm_yy_from_original_text(self):
        invoice_date = svc._extract_invoice_date_from_text(HENRY_SCHEIN_FIXTURE)
        self.assertEqual(invoice_date, "2026-04-22")

    def test_extract_invoice_date_with_longer_header_gap(self):
        invoice_date = svc._extract_invoice_date_from_text(HENRY_SCHEIN_FULL_DATE_FIXTURE)
        self.assertEqual(invoice_date, "2026-04-22")

    def test_text_total_does_not_override_tax_summary(self):
        summary = svc._extract_tax_summary_from_text(HENRY_SCHEIN_FULL_DATE_FIXTURE)
        self.assertTrue(summary.get("found"))
        base, vat, total, rate, source = svc._apply_tax_summary_override(
            HENRY_SCHEIN_FULL_DATE_FIXTURE,
            63.67,
            13.37,
            77.67,
            None,
            summary,
        )
        self.assertEqual(source, "regex_tax_summary")
        text_amounts = svc._extract_amounts_from_text(HENRY_SCHEIN_FULL_DATE_FIXTURE)
        if text_amounts.get("total") is not None and source != "regex_tax_summary":
            if total is None or text_amounts["total"] <= (total + 0.02):
                total = text_amounts["total"]
        self.assertAlmostEqual(base, 77.67, places=2)
        self.assertAlmostEqual(vat, 14.77, places=2)
        self.assertAlmostEqual(total, 92.44, places=2)

    def test_extract_multivat_tax_summary_from_text(self):
        summary = svc._extract_tax_summary_from_text(CANTABRIA_MULTIVAT_FIXTURE)
        self.assertTrue(summary.get("found"))
        self.assertAlmostEqual(summary.get("base_amount"), 573.71, places=2)
        self.assertAlmostEqual(summary.get("vat_amount"), 108.24, places=2)
        self.assertAlmostEqual(summary.get("total_amount"), 681.95, places=2)
        self.assertEqual(len(summary.get("breakdown") or []), 2)
        self.assertEqual(sorted(line["rate"] for line in summary["breakdown"]), [10.0, 21.0])

    def test_breakdown_wins_over_partial_tax_summary(self):
        normalized = svc.normalize_and_validate_amounts(
            {
                "analysis_status": "ok",
                "base_amount": 111.35,
                "vat_amount": 11.13,
                "total_amount": 122.48,
                "vat_rate": 10.0,
                "vat_breakdown": [
                    {"base": 111.35, "vat_amount": 11.14, "rate": None},
                    {"base": 462.36, "vat_amount": 97.10, "rate": None},
                ],
                "amount_source": "regex_tax_summary",
            }
        )
        self.assertEqual(normalized["amount_source"], "breakdown")
        self.assertAlmostEqual(normalized["base_amount"], 573.71, places=2)
        self.assertAlmostEqual(normalized["vat_amount"], 108.24, places=2)
        self.assertAlmostEqual(normalized["total_amount"], 681.95, places=2)
        self.assertEqual(len(normalized["vat_breakdown"]), 2)

    def test_corrected_amounts_override_raw_totals_payload(self):
        normalized = svc.normalize_and_validate_amounts(
            {
                "analysis_status": "ok",
                "base_amount": 573.71,
                "vat_amount": 108.24,
                "total_amount": 681.95,
                "vat_rate": None,
                "vat_breakdown": [
                    {"base": 111.35, "vat_amount": 11.14, "rate": 10.0, "total": 122.49},
                    {"base": 462.36, "vat_amount": 97.10, "rate": 21.0, "total": 559.46},
                ],
                "totals": {"base": 573.71, "vat": 97.10, "total": 681.95},
                "amount_source": "regex_tax_summary",
            }
        )
        self.assertEqual(normalized["analysis_status"], "ok")
        self.assertEqual(normalized["amount_source"], "regex_tax_summary")
        self.assertAlmostEqual(normalized["base_amount"], 573.71, places=2)
        self.assertAlmostEqual(normalized["vat_amount"], 108.24, places=2)
        self.assertAlmostEqual(normalized["total_amount"], 681.95, places=2)
        self.assertEqual(len(normalized["vat_breakdown"]), 2)

    def test_tax_summary_rescue_recovers_partial_result(self):
        summary = svc._extract_tax_summary_from_text(CANTABRIA_MULTIVAT_FIXTURE)
        rescued = svc._recover_from_tax_summary_if_needed(
            "partial",
            None,
            None,
            None,
            None,
            [],
            "fallback",
            summary,
        )
        self.assertEqual(rescued["analysis_status"], "ok")
        self.assertEqual(rescued["amount_source"], "regex_tax_summary")
        self.assertAlmostEqual(rescued["base_amount"], 573.71, places=2)
        self.assertAlmostEqual(rescued["vat_amount"], 108.24, places=2)
        self.assertAlmostEqual(rescued["total_amount"], 681.95, places=2)
        self.assertEqual(len(rescued["vat_breakdown"]), 2)

    def test_tax_summary_is_forced_over_fallback_result(self):
        summary = svc._extract_tax_summary_from_text(CANTABRIA_MULTIVAT_FIXTURE)
        forced = svc._force_tax_summary_result_if_available(
            "partial",
            None,
            None,
            None,
            None,
            [],
            "fallback",
            False,
            summary,
        )
        self.assertEqual(forced["analysis_status"], "ok")
        self.assertEqual(forced["amount_source"], "regex_tax_summary")
        self.assertAlmostEqual(forced["base_amount"], 573.71, places=2)
        self.assertAlmostEqual(forced["vat_amount"], 108.24, places=2)
        self.assertAlmostEqual(forced["total_amount"], 681.95, places=2)
        self.assertEqual(len(forced["vat_breakdown"]), 2)

    def test_extract_vertical_tax_summary_from_text(self):
        summary = svc._extract_tax_summary_from_text(YSONUT_FIXTURE)
        self.assertTrue(summary.get("found"))
        self.assertAlmostEqual(summary.get("base_amount"), 169.47, places=2)
        self.assertAlmostEqual(summary.get("vat_amount"), 16.95, places=2)
        self.assertAlmostEqual(summary.get("total_amount"), 186.42, places=2)
        self.assertEqual(len(summary.get("breakdown") or []), 1)
        self.assertAlmostEqual(summary["breakdown"][0]["rate"], 10.0, places=2)

    def test_extract_multiple_payment_dates_from_vertical_schedule(self):
        payment_dates = svc._find_payment_dates_by_keywords(YSONUT_FIXTURE, "2026-04-30")
        self.assertEqual(payment_dates, ["2026-05-31", "2026-06-15"])

    def test_ignore_legal_text_days_outside_payment_context(self):
        payment_dates = svc._find_payment_dates_by_keywords(YSONUT_LONG_FIXTURE, "2026-04-30")
        self.assertEqual(payment_dates, ["2026-05-31", "2026-06-15"])

    def test_ignore_footer_dates_after_single_due_date_block(self):
        payment_dates = svc._find_payment_dates_by_keywords(
            CANTABRIA_SINGLE_DUE_WITH_FOOTER_DATE_FIXTURE,
            "2026-05-08",
        )
        self.assertEqual(payment_dates, ["2026-07-07"])

    def test_explicit_due_date_beats_transaction_dates(self):
        payment_dates, payment_terms_days = svc._resolve_payment_schedule(
            RAIOLA_FIXTURE,
            "2026-05-11",
            ["2026-05-16"],
            None,
            5,
        )
        self.assertEqual(payment_terms_days, 5)
        self.assertEqual(payment_dates, ["2026-05-16"])

    def test_text_total_does_not_override_coherent_llm_total(self):
        normalized = svc.normalize_and_validate_amounts(
            {
                "analysis_status": "ok",
                "base_amount": 98.95,
                "vat_amount": 20.78,
                "total_amount": 119.73,
                "vat_rate": 21.0,
                "vat_breakdown": [{"rate": 21.0, "base": 98.95, "vat_amount": 20.78, "total": 119.73}],
                "totals": {"base": 98.95, "vat": 20.78, "total": 119.73},
                "amount_source": "llm",
            }
        )
        self.assertEqual(normalized["analysis_status"], "ok")
        self.assertAlmostEqual(normalized["total_amount"], 119.73, places=2)

    def test_explicit_payment_schedule_wins_over_zero_terms_days(self):
        payment_dates, payment_terms_days = svc._resolve_payment_schedule(
            YSONUT_LONG_FIXTURE,
            "2026-04-30",
            [],
            None,
            0,
        )
        self.assertEqual(payment_terms_days, None)
        self.assertEqual(payment_dates, ["2026-05-31", "2026-06-15"])

    def test_extract_supplier_ignores_legal_footer_text(self):
        supplier = svc._extract_supplier_from_text(
            SKINTECH_FIXTURE,
            ["KALOS HEALTH AND BEAUTY S.L."],
        )
        self.assertEqual(supplier, "Skin Tech Pharma Group, SLU")

    def test_strip_inline_tax_id_handles_dotted_cif_labels(self):
        cleaned = svc._strip_inline_tax_id("C.I.F. A26106013 -")
        self.assertEqual(cleaned, "")

    def test_legal_form_detection_does_not_match_plain_text(self):
        self.assertFalse(svc.has_legal_form("DESCUENTO EN CUOTA SERVICIO AMPLIACION"))

    def test_person_detection_rejects_domains_and_locations(self):
        self.assertFalse(svc.looks_like_person("www.verisure.es"))
        self.assertFalse(svc.looks_like_person("VALENCIA (VALENCIA)"))

    def test_metadata_detection_rejects_pharmacy_code_lines(self):
        self.assertTrue(svc._looks_like_metadata("Cod.Farmacia: 38"))

    def test_extract_supplier_from_multiline_legal_footer(self):
        supplier = svc._extract_supplier_from_text(
            VERISURE_FIXTURE,
            ["KALOS HEALTH AND BEAUTY S.L."],
        )
        self.assertEqual(supplier, "Securitas Direct España S.A.U")

    def test_address_line_is_not_accepted_as_supplier(self):
        self.assertFalse(
            svc._is_valid_supplier(
                "CALLE CERVANTES 25 BJ DRCHA",
                ["KALOS HEALTH AND BEAUTY S.L."],
                VERISURE_FIXTURE,
                require_tax_id=False,
            )
        )

    def test_legal_form_supplier_is_valid_without_nearby_tax_id(self):
        self.assertTrue(
            svc._is_valid_supplier(
                "Securitas Direct España S.A.U",
                ["KALOS HEALTH AND BEAUTY S.L."],
                VERISURE_FIXTURE,
                require_tax_id=True,
            )
        )

    def test_extract_amounts_prefers_total_factura_over_table_header(self):
        amounts = svc._extract_amounts_from_text(VERISURE_FIXTURE)
        self.assertAlmostEqual(amounts["base"], 60.51, places=2)
        self.assertAlmostEqual(amounts["total"], 73.22, places=2)

    def test_text_amount_fallbacks_replace_incoherent_llm_total(self):
        base_amount, vat_amount, total_amount, amount_source = svc._apply_text_amount_fallbacks(
            VERISURE_FIXTURE,
            60.51,
            12.71,
            76.85,
            "llm",
        )
        self.assertAlmostEqual(base_amount, 60.51, places=2)
        self.assertAlmostEqual(vat_amount, 12.71, places=2)
        self.assertAlmostEqual(total_amount, 73.22, places=2)
        self.assertEqual(amount_source, "text_total")

    def test_installment_lines_do_not_count_as_withholding(self):
        withholding_amount = svc._extract_explicit_withholding_amount_from_text(VERISURE_FIXTURE)
        self.assertIsNone(withholding_amount)

    def test_extracts_withholding_from_dotted_irpf_rent_columns(self):
        withholding_amount = svc._extract_explicit_withholding_amount_from_text(
            RENT_INVOICE_IRPF_COLUMNS_FIXTURE
        )
        self.assertEqual(withholding_amount, 57.0)

    def test_rent_totals_remain_valid_after_withholding(self):
        normalized = svc.normalize_and_validate_amounts(
            {
                "analysis_status": "ok",
                "base_amount": 300.0,
                "vat_amount": 63.0,
                "total_amount": 306.0,
                "withholding_amount": 57.0,
                "vat_breakdown": [{"base": 300.0, "vat_amount": 63.0, "rate": 21}],
            }
        )
        self.assertEqual(normalized["analysis_status"], "ok")
        self.assertEqual(normalized["base_amount"], 300.0)
        self.assertEqual(normalized["vat_amount"], 63.0)
        self.assertEqual(normalized["withholding_amount"], 57.0)
        self.assertEqual(normalized["total_amount"], 306.0)

    def test_detects_explicit_vat_exemption_without_assuming_standard_rate(self):
        exempt_base = svc._extract_explicit_vat_exemption_amount_from_text(
            VAT_EXEMPT_CONFIRMING_FIXTURE
        )
        self.assertEqual(exempt_base, 16.95)
        self.assertTrue(svc._validate_math(exempt_base, 0.0, 16.95)["is_consistent"])

    def test_extracts_irpf_withholding_from_professional_supplier_invoice(self):
        withholding_amount = svc._extract_explicit_withholding_amount_from_text(
            PROFESSIONAL_INVOICE_WITHHOLDING_FIXTURE
        )
        self.assertEqual(withholding_amount, 339.62)
        self.assertTrue(
            svc._validate_math(2264.15, 475.47, 2400.0, withholding_amount)["is_consistent"]
        )

    def test_autonomo_name_with_nearby_tax_id_is_valid_supplier(self):
        self.assertTrue(
            svc._is_valid_supplier(
                "CRISTINA SIMÓ BESALDUCH",
                ["KALOS HEALTH AND BEAUTY SL"],
                AUTONOMO_PHARMACY_FIXTURE,
                require_tax_id=True,
            )
        )

    def test_address_line_with_customer_number_is_not_valid_supplier(self):
        self.assertFalse(
            svc._is_valid_supplier(
                "C/CERVANTES Nº25 BAJO",
                ["KALOS HEALTH AND BEAUTY SL"],
                AUTONOMO_PHARMACY_FIXTURE,
                require_tax_id=False,
            )
        )

    def test_extract_supplier_prefers_autonomo_over_customer_address(self):
        supplier = svc._extract_supplier_from_text(
            AUTONOMO_PHARMACY_FIXTURE,
            ["KALOS HEALTH AND BEAUTY SL"],
        )
        self.assertEqual(supplier, "CRISTINA SIMÓ BESALDUCH")

    def test_extract_supplier_prioritizes_legal_entity_over_autonomo(self):
        supplier = svc._extract_supplier_from_text(
            MIXED_SUPPLIER_PRIORITY_FIXTURE,
            ["KALOS HEALTH AND BEAUTY SL"],
        )
        self.assertEqual(supplier, "SUMINISTROS CLÍNICOS LEVANTE, S.L.")

    def test_single_word_legal_entity_is_valid_supplier(self):
        self.assertTrue(
            svc._is_valid_supplier(
                "MEDIDERMA, S.L.U.",
                ["KALOS HEALTH AND BEAUTY S.L."],
                MEDIDERMA_CLIENT_TYPO_FIXTURE,
                require_tax_id=False,
            )
        )

    def test_client_context_typo_is_not_valid_supplier(self):
        self.assertFalse(
            svc._is_valid_supplier(
                "KARLOS HEALTH AND BEAUTY S.L.",
                ["KALOS HEALTH AND BEAUTY S.L."],
                MEDIDERMA_CLIENT_TYPO_FIXTURE,
                require_tax_id=False,
            )
        )

    def test_extract_supplier_ignores_client_typo_block(self):
        supplier = svc._extract_supplier_from_text(
            MEDIDERMA_CLIENT_TYPO_FIXTURE,
            ["KALOS HEALTH AND BEAUTY S.L."],
        )
        self.assertEqual(supplier, "MEDIDERMA, S.L.U")

    def test_tax_summary_ignores_percentages_inline_with_base_labels(self):
        summary = svc._extract_tax_summary_from_text(CONSENUR_FIXTURE)
        self.assertTrue(summary.get("found"))
        self.assertAlmostEqual(summary.get("base_amount"), 92.92, places=2)
        self.assertAlmostEqual(summary.get("vat_amount"), 9.29, places=2)
        self.assertAlmostEqual(summary.get("total_amount"), 102.21, places=2)
        self.assertAlmostEqual(summary.get("vat_rate"), 10.0, places=2)

    def test_text_amount_extraction_ignores_inline_percentage_amounts(self):
        amounts = svc._extract_amounts_from_text(CONSENUR_FIXTURE)
        self.assertAlmostEqual(amounts["base"], 92.92, places=2)
        self.assertAlmostEqual(amounts["vat"], 9.29, places=2)
        self.assertAlmostEqual(amounts["total"], 102.21, places=2)

    def test_nonstandard_breakdown_does_not_override_coherent_totals(self):
        normalized = svc.normalize_and_validate_amounts(
            {
                "analysis_status": "ok",
                "base_amount": 1262.80,
                "vat_amount": 0.0,
                "total_amount": 1262.80,
                "vat_breakdown": [
                    {"base": 115.44, "vat_amount": 66.96, "rate": None},
                    {"base": 160.00, "vat_amount": 92.80, "rate": None},
                ],
                "amount_source": "llm",
            }
        )
        self.assertEqual(normalized["amount_source"], "llm")
        self.assertAlmostEqual(normalized["base_amount"], 1262.80, places=2)
        self.assertAlmostEqual(normalized["vat_amount"], 0.0, places=2)
        self.assertAlmostEqual(normalized["total_amount"], 1262.80, places=2)
        self.assertEqual(normalized["vat_breakdown"], [])


if __name__ == "__main__":
    unittest.main()
