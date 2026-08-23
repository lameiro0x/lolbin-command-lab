#!/usr/bin/env bash
# Launcher interactivo de laboratorio local.
# Taxonomía y metadatos integrados desde excel_limpio_estructural.xlsx.
# No ejecuta muestras/IOCs originales: las validaciones son benignas y locales.
set -u

SCRIPT_NAME="${0##*/}"
PAGE_SIZE=15

SUMMARY_TSV=$(cat <<'__SUMMARY_TSV__'
Categoria	Categoria_ID	Categoria_Nombre	Num_Grupos	Num_Registros	Pct_Registros	Dominios_Distintos	Registros_Host_IP	Grupo_Mayor	Tam_Grupo_Mayor	Grupos_Unitarios	Captcha_Si	Fecha_Primera	Fecha_Ultima	Longitud_Media	Grupos_Descarga_Remota	Grupos_Codificacion_Base64	Grupos_Ofuscacion	Grupos_Ejecucion_Oculta	Grupos_Persistencia	Filas_Mal_Parseadas
00 - Sin contenido	00	Sin contenido	1	261	6.02	0	0	2	261	0	0			0	0	0	0	0	0	261
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	23	1031	23.76	987	44	1	651	7	10	2025-08-22	2026-07-27	98	23	0	1	0	0	0
02 - cmd + WebDAV (net use)	02	cmd + WebDAV (net use)	4	280	6.45	279	1	3	153	0	128	2026-03-13	2026-07-29	124	4	0	0	0	0	0
03 - conhost --headless	03	conhost --headless	17	48	1.11	47	1	20	29	15	14	2025-08-20	2026-07-29	208	9	3	8	15	0	0
04 - mshta	04	mshta	66	81	1.87	81	0	48	7	60	7	2025-08-24	2026-05-20	51	66	0	0	0	0	0
05 - rundll32	05	rundll32	4	4	0.09	4	0	939	1	4	3	2026-03-04	2026-05-02	91	4	0	0	0	0	0
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	22	24	0.55	24	0	91	3	21	24	2026-01-09	2026-06-19	151	2	0	0	0	0	0
07 - pcalua	07	pcalua	4	4	0.09	4	0	241	1	4	4	2026-07-18	2026-07-22	229	4	0	2	0	0	0
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	515	914	21.06	908	6	8	82	430	394	2025-08-17	2026-07-30	188	284	68	254	329	1	0
09 - cmd.exe	09	cmd.exe	68	633	14.59	499	133	4	151	42	35	2025-08-20	2026-07-22	168	58	1	7	26	0	1
10 - script PowerShell embebido	10	script PowerShell embebido	722	998	23	995	1	6	121	708	349	2025-08-20	2026-07-28	1616	139	244	592	126	0	2
11 - shell Unix (bash/ssh/curl)	11	shell Unix (bash/ssh/curl)	6	18	0.41	18	0	47	8	3	16	2025-08-21	2026-06-30	98	5	1	0	0	0	0
12 - solo URL / dominio	12	solo URL / dominio	10	13	0.3	13	0	74	4	9	6	2025-11-11	2026-07-29	51	9	0	0	0	0	0
13 - echo	13	echo	5	5	0.12	5	0	1169	1	5	2	2026-01-04	2026-03-15	150	0	4	0	0	0	0
99 - otro	99	otro	16	25	0.58	22	1	41	9	14	13	2025-08-21	2026-07-24	40	4	0	3	3	0	2
ZZ - fila vacia	ZZ	fila vacia	0	72	0	0	0		0	0	0			0	0	0	0	0	0	0
__SUMMARY_TSV__
)

GROUP_TSV=$(cat <<'__GROUP_TSV__'
Categoria	Categoria_ID	Categoria_Nombre	Grupo	Repeticiones	Num_Registros	Dominios_Distintos	Registros_Host_IP	Fecha_Primera	Fecha_Ultima	Captcha_Si	Captcha_No	Longitud_Media	Longitud_Max	Multilinea	Descarga_Remota	Codificacion_Base64	Ofuscacion	Ejecucion_Oculta	Persistencia	Filas_Mal_Parseadas	Grupo_Hash
00 - Sin contenido	00	Sin contenido	2	261	261	0	0			0	0	0	0	No	No	No	No	No	No	261	
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	1	651	651	646	5	2025-11-15	2026-02-19	0	651	113	113	No	Si	No	No	No	No	0	bf4e30a8d948
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	9	78	78	62	16	2025-12-28	2026-02-28	0	78	47	47	No	Si	No	No	No	No	0	1485de84ed93
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	12	54	54	54	0	2025-11-22	2025-12-08	1	53	99	99	No	Si	No	No	No	No	0	b6ee2694f9c7
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	13	44	44	43	1	2025-12-03	2026-02-06	0	44	85	85	No	Si	No	No	No	No	0	0efa15047713
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	17	35	35	34	1	2025-11-21	2026-01-27	0	35	87	87	No	Si	No	No	No	No	0	378549f8b164
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	18	34	34	19	15	2025-12-29	2026-02-06	0	34	43	43	No	Si	No	No	No	No	0	e980d8939d14
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	23	24	24	24	0	2025-10-28	2026-02-02	1	23	94	94	No	Si	No	No	No	No	0	482beeddffb1
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	24	23	23	23	0	2025-11-03	2025-12-08	0	23	112	112	No	Si	No	No	No	No	0	84b1dca2196c
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	27	21	21	21	0	2026-04-04	2026-04-14	0	21	35	35	No	Si	No	No	No	No	0	9efe0febcc46
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	29	21	21	21	0	2025-11-28	2025-12-18	2	19	92	92	No	Si	No	No	No	No	0	ccf74a02a376
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	33	18	18	13	5	2025-12-31	2026-02-06	0	18	34	34	No	Si	No	No	No	No	0	e120944c6a6c
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	54	6	6	6	0	2025-12-02	2025-12-09	0	6	47	47	No	Si	No	No	No	No	0	612485e5deb0
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	60	5	5	5	0	2026-05-31	2026-06-17	3	2	95	95	No	Si	No	No	No	No	0	5252024cb7f9
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	78	4	4	4	0	2025-11-04	2025-12-12	0	4	90	90	No	Si	No	No	No	No	0	65e56d72764d
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	84	3	3	3	0	2026-05-16	2026-05-20	0	3	43	43	No	Si	No	No	No	No	0	ebf6ea1baed5
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	94	3	3	3	0	2025-10-25	2025-11-21	0	3	45	45	No	Si	No	No	No	No	0	9a6cde160fd5
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	204	1	1	1	0	2026-07-27	2026-07-27	1	0	105	105	No	Si	No	No	No	No	0	fa50260c5c05
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	521	1	1	1	0	2026-06-19	2026-06-19	0	1	86	86	No	Si	No	No	No	No	0	6726b4ca402e
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	1124	1	1	1	0	2026-03-26	2026-03-26	1	0	85	85	No	Si	No	No	No	No	0	5669b1516618
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	1301	1	1	0	1	2025-12-31	2025-12-31	0	1	46	46	No	Si	No	No	No	No	0	3d09b2b7f95a
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	1317	1	1	1	0	2025-12-08	2025-12-08	0	1	51	51	No	Si	No	No	No	No	0	cf3810efe547
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	1358	1	1	1	0	2025-11-03	2025-11-03	0	1	116	116	No	Si	No	No	No	No	0	1bc41d8ce482
01 - msiexec (MSI remoto)	01	msiexec (MSI remoto)	1470	1	1	1	0	2025-08-22	2025-08-22	1	0	130	130	No	Si	No	Si	No	No	0	0ab6ec8bb79b
02 - cmd + WebDAV (net use)	02	cmd + WebDAV (net use)	3	153	153	152	1	2026-03-21	2026-04-11	2	151	144	144	No	Si	No	No	No	No	0	9dfb113fa818
02 - cmd + WebDAV (net use)	02	cmd + WebDAV (net use)	7	118	118	118	0	2026-03-13	2026-07-29	117	1	101	101	No	Si	No	No	No	No	0	a96ccd1767b4
02 - cmd + WebDAV (net use)	02	cmd + WebDAV (net use)	59	5	5	5	0	2026-06-17	2026-06-18	5	0	65	65	No	Si	No	No	No	No	0	d89c926b6e9f
02 - cmd + WebDAV (net use)	02	cmd + WebDAV (net use)	73	4	4	4	0	2026-04-02	2026-06-17	4	0	101	101	No	Si	No	No	No	No	0	b64c2cd9c97c
03 - conhost --headless	03	conhost --headless	20	29	29	29	0	2026-07-01	2026-07-17	6	23	220	220	No	Si	No	No	Si	No	0	c23a27a99f3e
03 - conhost --headless	03	conhost --headless	72	4	4	4	0	2026-07-20	2026-07-21	0	4	220	220	No	Si	No	No	Si	No	0	16ec7775b3c6
03 - conhost --headless	03	conhost --headless	163	1	1	1	0	2026-07-29	2026-07-29	1	0	166	166	No	Si	No	Si	Si	No	0	f2bfaee8f20a
03 - conhost --headless	03	conhost --headless	168	1	1	1	0	2026-07-28	2026-07-28	1	0	155	155	No	Si	No	Si	Si	No	0	cd5052c9b215
03 - conhost --headless	03	conhost --headless	170	1	1	1	0	2026-07-28	2026-07-28	1	0	154	154	No	Si	No	Si	Si	No	0	b709d426c080
03 - conhost --headless	03	conhost --headless	232	1	1	1	0	2026-07-26	2026-07-26	1	0	181	181	No	Si	No	Si	Si	No	0	57f13ffedaac
03 - conhost --headless	03	conhost --headless	235	1	1	1	0	2026-07-25	2026-07-25	1	0	165	165	No	Si	No	Si	Si	No	0	6b4b7c3e5016
03 - conhost --headless	03	conhost --headless	388	1	1	1	0	2026-06-24	2026-06-24	1	0	175	175	No	No	No	Si	Si	No	0	617129f9359b
03 - conhost --headless	03	conhost --headless	516	1	1	1	0	2026-06-21	2026-06-21	0	1	132	132	No	No	No	Si	Si	No	0	08395728f116
03 - conhost --headless	03	conhost --headless	517	1	1	1	0	2026-06-20	2026-06-20	0	1	125	125	No	No	No	Si	Si	No	0	aebaf0c6a2d8
03 - conhost --headless	03	conhost --headless	556	1	1	1	0	2026-06-08	2026-06-08	0	1	161	161	No	No	No	No	Si	No	0	0e219690ce29
03 - conhost --headless	03	conhost --headless	602	1	1	1	0	2026-05-31	2026-05-31	1	0	154	154	No	Si	No	No	Si	No	0	09904538a191
03 - conhost --headless	03	conhost --headless	663	1	1	1	0	2026-05-18	2026-05-18	0	1	143	143	No	Si	No	No	Si	No	0	806acb99f7b3
03 - conhost --headless	03	conhost --headless	1255	1	1	1	0	2026-01-27	2026-01-27	0	1	105	105	No	No	No	No	Si	No	0	207dc96a97f0
03 - conhost --headless	03	conhost --headless	1316	1	1	0	1	2025-12-08	2025-12-08	0	1	241	241	No	No	Si	No	No	No	0	cbb63bbd22e5
03 - conhost --headless	03	conhost --headless	1371	1	1	1	0	2025-10-15	2025-10-15	1	0	241	241	No	No	Si	No	No	No	0	da382488eaef
03 - conhost --headless	03	conhost --headless	1481	1	1	1	0	2025-08-20	2025-08-20	0	1	421	421	No	No	Si	No	Si	No	0	847a19faf7fb
04 - mshta	04	mshta	48	7	7	7	0	2025-11-13	2025-12-11	0	7	37	37	No	Si	No	No	No	No	0	9d91d2dca2e7
04 - mshta	04	mshta	65	5	5	5	0	2025-11-14	2025-11-26	0	5	33	33	No	Si	No	No	No	No	0	cfec9f4a9818
04 - mshta	04	mshta	87	3	3	3	0	2026-03-02	2026-03-04	0	3	36	36	No	Si	No	No	No	No	0	00f89a10beb7
04 - mshta	04	mshta	130	2	2	2	0	2026-03-03	2026-03-04	0	2	37	37	No	Si	No	No	No	No	0	6b7af18a9974
04 - mshta	04	mshta	131	2	2	2	0	2026-02-28	2026-02-28	2	0	45	45	No	Si	No	No	No	No	0	5a81cce59523
04 - mshta	04	mshta	133	2	2	2	0	2026-02-03	2026-02-03	0	2	27	27	No	Si	No	No	No	No	0	6e147d859c67
04 - mshta	04	mshta	656	1	1	1	0	2026-05-20	2026-05-20	0	1	103	103	No	Si	No	No	No	No	0	d373b03eeaa1
04 - mshta	04	mshta	754	1	1	1	0	2026-05-11	2026-05-11	0	1	24	24	No	Si	No	No	No	No	0	18894f974c94
04 - mshta	04	mshta	1223	1	1	1	0	2026-03-09	2026-03-09	0	1	41	41	No	Si	No	No	No	No	0	75ffaa2287e2
04 - mshta	04	mshta	1224	1	1	1	0	2026-03-09	2026-03-09	0	1	41	41	No	Si	No	No	No	No	0	12d8865b3902
04 - mshta	04	mshta	1237	1	1	1	0	2026-03-02	2026-03-02	1	0	38	38	No	Si	No	No	No	No	0	a13c164afeb5
04 - mshta	04	mshta	1238	1	1	1	0	2026-03-01	2026-03-01	0	1	41	41	No	Si	No	No	No	No	0	9b4df90d046a
04 - mshta	04	mshta	1239	1	1	1	0	2026-02-28	2026-02-28	0	1	41	41	No	Si	No	No	No	No	0	9a68486c4dd5
04 - mshta	04	mshta	1248	1	1	1	0	2026-02-18	2026-02-18	0	1	134	134	No	Si	No	No	No	No	0	a5fd38e2eb1f
04 - mshta	04	mshta	1251	1	1	1	0	2026-02-03	2026-02-03	0	1	28	28	No	Si	No	No	No	No	0	070c5736a9cd
04 - mshta	04	mshta	1258	1	1	1	0	2026-01-21	2026-01-21	0	1	35	35	No	Si	No	No	No	No	0	313b9d3777d4
04 - mshta	04	mshta	1259	1	1	1	0	2026-01-21	2026-01-21	0	1	36	36	No	Si	No	No	No	No	0	6574c7549660
04 - mshta	04	mshta	1263	1	1	1	0	2026-01-17	2026-01-17	1	0	51	51	No	Si	No	No	No	No	0	1e9f16aeeb82
04 - mshta	04	mshta	1297	1	1	1	0	2026-01-04	2026-01-04	1	0	84	84	No	Si	No	No	No	No	0	804af4b12478
04 - mshta	04	mshta	1304	1	1	1	0	2025-12-29	2025-12-29	1	0	131	131	No	Si	No	No	No	No	0	8d26f4037a37
04 - mshta	04	mshta	1318	1	1	1	0	2025-12-08	2025-12-08	0	1	39	39	No	Si	No	No	No	No	0	273f68ad5a52
04 - mshta	04	mshta	1330	1	1	1	0	2025-11-29	2025-11-29	0	1	37	37	No	Si	No	No	No	No	0	04138832abb9
04 - mshta	04	mshta	1336	1	1	1	0	2025-11-23	2025-11-23	0	1	40	40	No	Si	No	No	No	No	0	53519ede3cdc
04 - mshta	04	mshta	1337	1	1	1	0	2025-11-23	2025-11-23	0	1	37	37	No	Si	No	No	No	No	0	fb8f8ef15874
04 - mshta	04	mshta	1352	1	1	1	0	2025-11-06	2025-11-06	0	1	204	204	No	Si	No	No	No	No	0	4a4cedca6019
04 - mshta	04	mshta	1356	1	1	1	0	2025-11-04	2025-11-04	0	1	93	93	No	Si	No	No	No	No	0	7f7f6803ded3
04 - mshta	04	mshta	1363	1	1	1	0	2025-10-28	2025-10-28	0	1	93	93	No	Si	No	No	No	No	0	a7d7f8997dd7
04 - mshta	04	mshta	1365	1	1	1	0	2025-10-24	2025-10-24	0	1	93	93	No	Si	No	No	No	No	0	6c54f4bcd502
04 - mshta	04	mshta	1376	1	1	1	0	2025-10-11	2025-10-11	0	1	45	45	No	Si	No	No	No	No	0	e76204bf334e
04 - mshta	04	mshta	1389	1	1	1	0	2025-10-02	2025-10-02	0	1	48	48	No	Si	No	No	No	No	0	6d23f4d8b463
04 - mshta	04	mshta	1390	1	1	1	0	2025-10-02	2025-10-02	0	1	118	118	No	Si	No	No	No	No	0	cd146ca9c0bc
04 - mshta	04	mshta	1392	1	1	1	0	2025-10-01	2025-10-01	0	1	47	47	No	Si	No	No	No	No	0	408814717518
04 - mshta	04	mshta	1402	1	1	1	0	2025-09-16	2025-09-16	1	0	34	34	No	Si	No	No	No	No	0	4094c932cb29
04 - mshta	04	mshta	1405	1	1	1	0	2025-09-12	2025-09-12	0	1	47	47	No	Si	No	No	No	No	0	c549be380ffd
04 - mshta	04	mshta	1411	1	1	1	0	2025-09-04	2025-09-04	0	1	51	51	No	Si	No	No	No	No	0	e9e07861be3e
04 - mshta	04	mshta	1412	1	1	1	0	2025-09-04	2025-09-04	0	1	50	50	No	Si	No	No	No	No	0	d43b3480eea5
04 - mshta	04	mshta	1414	1	1	1	0	2025-09-04	2025-09-04	0	1	51	51	No	Si	No	No	No	No	0	0a78a0c61596
04 - mshta	04	mshta	1416	1	1	1	0	2025-09-03	2025-09-03	0	1	50	50	No	Si	No	No	No	No	0	80423b2b4771
04 - mshta	04	mshta	1417	1	1	1	0	2025-09-02	2025-09-02	0	1	50	50	No	Si	No	No	No	No	0	f68062f4a4dd
04 - mshta	04	mshta	1419	1	1	1	0	2025-09-02	2025-09-02	0	1	50	50	No	Si	No	No	No	No	0	543519be67d4
04 - mshta	04	mshta	1420	1	1	1	0	2025-09-02	2025-09-02	0	1	50	50	No	Si	No	No	No	No	0	a3eb2e5690a9
04 - mshta	04	mshta	1423	1	1	1	0	2025-09-01	2025-09-01	0	1	49	49	No	Si	No	No	No	No	0	25e8d69747e1
04 - mshta	04	mshta	1424	1	1	1	0	2025-09-01	2025-09-01	0	1	49	49	No	Si	No	No	No	No	0	fe7ccaab0544
04 - mshta	04	mshta	1425	1	1	1	0	2025-09-01	2025-09-01	0	1	50	50	No	Si	No	No	No	No	0	2b64db8afc84
04 - mshta	04	mshta	1428	1	1	1	0	2025-08-29	2025-08-29	0	1	52	52	No	Si	No	No	No	No	0	b3bca52bb83c
04 - mshta	04	mshta	1434	1	1	1	0	2025-08-28	2025-08-28	0	1	49	49	No	Si	No	No	No	No	0	81e0dcdcf6e8
04 - mshta	04	mshta	1435	1	1	1	0	2025-08-28	2025-08-28	0	1	49	49	No	Si	No	No	No	No	0	a92a49e5eb82
04 - mshta	04	mshta	1437	1	1	1	0	2025-08-28	2025-08-28	0	1	49	49	No	Si	No	No	No	No	0	d780b696043a
04 - mshta	04	mshta	1438	1	1	1	0	2025-08-27	2025-08-27	0	1	49	49	No	Si	No	No	No	No	0	3fa31081da59
04 - mshta	04	mshta	1439	1	1	1	0	2025-08-27	2025-08-27	0	1	49	49	No	Si	No	No	No	No	0	4e4176c8c2b6
04 - mshta	04	mshta	1440	1	1	1	0	2025-08-27	2025-08-27	0	1	49	49	No	Si	No	No	No	No	0	d579990a71ea
04 - mshta	04	mshta	1441	1	1	1	0	2025-08-27	2025-08-27	0	1	49	49	No	Si	No	No	No	No	0	375754dd6128
04 - mshta	04	mshta	1442	1	1	1	0	2025-08-27	2025-08-27	0	1	49	49	No	Si	No	No	No	No	0	4bd4d47c94db
04 - mshta	04	mshta	1443	1	1	1	0	2025-08-27	2025-08-27	0	1	49	49	No	Si	No	No	No	No	0	831624ab9d48
04 - mshta	04	mshta	1445	1	1	1	0	2025-08-26	2025-08-26	0	1	60	60	No	Si	No	No	No	No	0	bc16d52f76c9
04 - mshta	04	mshta	1446	1	1	1	0	2025-08-26	2025-08-26	0	1	50	50	No	Si	No	No	No	No	0	b955fb494c54
04 - mshta	04	mshta	1447	1	1	1	0	2025-08-26	2025-08-26	0	1	49	49	No	Si	No	No	No	No	0	ad1846e6654e
04 - mshta	04	mshta	1449	1	1	1	0	2025-08-26	2025-08-26	0	1	49	49	No	Si	No	No	No	No	0	ad76d1676f08
04 - mshta	04	mshta	1450	1	1	1	0	2025-08-26	2025-08-26	0	1	49	49	No	Si	No	No	No	No	0	d830e3e7e2a4
04 - mshta	04	mshta	1451	1	1	1	0	2025-08-26	2025-08-26	0	1	49	49	No	Si	No	No	No	No	0	4d4df3b108c9
04 - mshta	04	mshta	1453	1	1	1	0	2025-08-26	2025-08-26	0	1	49	49	No	Si	No	No	No	No	0	e0dd00b224d6
04 - mshta	04	mshta	1454	1	1	1	0	2025-08-25	2025-08-25	0	1	49	49	No	Si	No	No	No	No	0	74f69775f373
04 - mshta	04	mshta	1457	1	1	1	0	2025-08-25	2025-08-25	0	1	49	49	No	Si	No	No	No	No	0	cafaade3c526
04 - mshta	04	mshta	1458	1	1	1	0	2025-08-25	2025-08-25	0	1	49	49	No	Si	No	No	No	No	0	8cbd2bcdd727
04 - mshta	04	mshta	1460	1	1	1	0	2025-08-25	2025-08-25	0	1	49	49	No	Si	No	No	No	No	0	abc6420a9acc
04 - mshta	04	mshta	1463	1	1	1	0	2025-08-24	2025-08-24	0	1	49	49	No	Si	No	No	No	No	0	ff99b473f05c
05 - rundll32	05	rundll32	939	1	1	1	0	2026-05-02	2026-05-02	1	0	100	100	No	Si	No	No	No	No	0	3e23f879fd4e
05 - rundll32	05	rundll32	1025	1	1	1	0	2026-04-14	2026-04-14	0	1	125	125	No	Si	No	No	No	No	0	8f8f9f5e4273
05 - rundll32	05	rundll32	1226	1	1	1	0	2026-03-08	2026-03-08	1	0	70	70	No	Si	No	No	No	No	0	079f0a08125b
05 - rundll32	05	rundll32	1236	1	1	1	0	2026-03-04	2026-03-04	1	0	68	68	No	Si	No	No	No	No	0	c8d84fcfac24
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	91	3	3	3	0	2026-01-12	2026-01-12	3	0	134	134	No	No	No	No	No	No	0	788820eae165
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	519	1	1	1	0	2026-06-19	2026-06-19	1	0	98	98	No	Si	No	No	No	No	0	2b9366c78b90
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1244	1	1	1	0	2026-02-22	2026-02-22	1	0	117	117	No	Si	No	No	No	No	0	5dc9554d492a
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1256	1	1	1	0	2026-01-23	2026-01-23	1	0	132	132	No	No	No	No	No	No	0	89dd0324c4aa
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1265	1	1	1	0	2026-01-16	2026-01-16	1	0	125	125	No	No	No	No	No	No	0	2958edd8eca8
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1266	1	1	1	0	2026-01-15	2026-01-15	1	0	114	114	No	No	No	No	No	No	0	a299a7228022
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1269	1	1	1	0	2026-01-12	2026-01-12	1	0	151	151	No	No	No	No	No	No	0	f20d4b53549c
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1270	1	1	1	0	2026-01-12	2026-01-12	1	0	161	161	No	No	No	No	No	No	0	f9be1595efd2
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1271	1	1	1	0	2026-01-11	2026-01-11	1	0	171	171	No	No	No	No	No	No	0	d17f302dde38
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1274	1	1	1	0	2026-01-10	2026-01-10	1	0	142	142	No	No	No	No	No	No	0	e83a27cbdfdb
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1276	1	1	1	0	2026-01-10	2026-01-10	1	0	169	169	No	No	No	No	No	No	0	05684d7658fc
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1277	1	1	1	0	2026-01-10	2026-01-10	1	0	169	169	No	No	No	No	No	No	0	8d29f8e794a2
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1278	1	1	1	0	2026-01-10	2026-01-10	1	0	169	169	No	No	No	No	No	No	0	13a5f242048a
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1279	1	1	1	0	2026-01-10	2026-01-10	1	0	169	169	No	No	No	No	No	No	0	71d6f148e892
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1280	1	1	1	0	2026-01-10	2026-01-10	1	0	169	169	No	No	No	No	No	No	0	ae7d538542d2
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1281	1	1	1	0	2026-01-10	2026-01-10	1	0	169	169	No	No	No	No	No	No	0	94e6e9dbc996
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1282	1	1	1	0	2026-01-10	2026-01-10	1	0	169	169	No	No	No	No	No	No	0	1ffcaa0813af
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1283	1	1	1	0	2026-01-10	2026-01-10	1	0	169	169	No	No	No	No	No	No	0	3ed39b52cfc3
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1284	1	1	1	0	2026-01-10	2026-01-10	1	0	169	169	No	No	No	No	No	No	0	962d483552a2
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1285	1	1	1	0	2026-01-10	2026-01-10	1	0	169	169	No	No	No	No	No	No	0	c2bbbdd98f52
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1286	1	1	1	0	2026-01-10	2026-01-10	1	0	169	169	No	No	No	No	No	No	0	c763c9d0f04f
06 - wscript / cscript / SyncAppvPublishingServer	06	wscript / cscript / SyncAppvPublishingServer	1287	1	1	1	0	2026-01-09	2026-01-09	1	0	159	159	No	No	No	No	No	No	0	2f29fbf20f60
07 - pcalua	07	pcalua	241	1	1	1	0	2026-07-22	2026-07-22	1	0	159	159	No	Si	No	Si	No	No	0	307a060f8bea
07 - pcalua	07	pcalua	244	1	1	1	0	2026-07-22	2026-07-22	1	0	255	255	No	Si	No	Si	No	No	0	d2de0d3778b3
07 - pcalua	07	pcalua	252	1	1	1	0	2026-07-19	2026-07-19	1	0	251	251	No	Si	No	No	No	No	0	b0f868d80a05
07 - pcalua	07	pcalua	254	1	1	1	0	2026-07-18	2026-07-18	1	0	251	251	No	Si	No	No	No	No	0	847abea5c2c2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	8	82	82	82	0	2025-08-20	2026-02-02	0	82	248	248	No	No	No	Si	Si	No	0	c6d5e2aa29b7
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	21	25	25	25	0	2026-02-17	2026-03-06	0	25	107	107	No	Si	No	Si	No	No	0	95ca565115dd
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	25	23	23	23	0	2025-08-17	2025-11-27	1	22	145	145	No	Si	No	No	No	No	0	1aa0c381bfa1
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	26	21	21	21	0	2026-07-19	2026-07-20	12	9	245	245	No	Si	No	Si	Si	No	0	6f6206ed8098
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	28	21	21	21	0	2026-02-02	2026-03-22	17	4	53	53	No	No	No	Si	No	No	0	3b0518a23e9d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	30	20	20	20	0	2026-03-19	2026-07-23	20	0	197	197	No	No	No	Si	No	No	0	ee8d2fc0511d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	31	20	20	20	0	2026-03-28	2026-03-30	0	20	209	209	No	Si	No	No	No	No	0	571edeefbc27
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	36	17	17	17	0	2026-06-29	2026-07-03	17	0	114	114	No	No	No	Si	No	No	0	7f471b28be05
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	39	10	10	10	0	2025-10-24	2026-04-04	10	0	111	111	No	Si	No	No	No	No	0	4ecb38cf8c13
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	40	10	10	10	0	2026-01-04	2026-01-09	10	0	177	177	No	No	Si	No	No	No	0	0a3d41915eec
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	42	9	9	9	0	2026-01-08	2026-01-30	8	1	53	53	No	No	No	Si	No	No	0	d2338b757963
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	44	9	9	9	0	2025-08-20	2025-12-08	0	9	141	141	No	Si	No	No	No	No	0	448ad1075cca
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	46	8	8	8	0	2026-03-22	2026-03-23	0	8	96	96	No	No	No	Si	Si	No	0	8044c2a13149
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	50	7	7	7	0	2025-08-20	2025-08-24	7	0	178	178	No	Si	No	Si	Si	No	0	3f117db2070c
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	51	6	6	6	0	2026-03-31	2026-04-05	6	0	248	248	No	No	Si	No	No	No	0	8391fd76e1e6
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	53	6	6	6	0	2026-03-09	2026-03-09	0	6	55	55	No	No	No	Si	Si	No	0	b665451263aa
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	55	6	6	6	0	2025-08-20	2025-10-09	0	6	142	142	No	Si	No	No	No	No	0	495c98a90705
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	56	6	6	6	0	2025-08-26	2025-08-27	6	0	165	165	No	Si	No	Si	Si	No	0	49f1ce24ed65
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	57	5	5	5	0	2026-07-26	2026-07-26	3	2	241	241	No	Si	No	Si	Si	No	0	cf2ab462669a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	58	5	5	5	0	2026-07-26	2026-07-26	1	4	233	233	No	Si	No	Si	No	No	0	d33fcc2549b1
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	61	5	5	5	0	2026-04-01	2026-04-07	5	0	98	98	No	No	No	Si	No	No	0	3705e37f270d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	63	5	5	5	0	2025-12-29	2025-12-29	2	3	55	55	No	No	No	Si	No	No	0	30ab2e2f2441
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	64	5	5	5	0	2025-12-02	2025-12-12	0	5	205	205	No	Si	No	Si	Si	No	0	66c9d4a0a074
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	66	5	5	5	0	2025-10-07	2025-10-08	5	0	198	198	No	Si	No	No	Si	No	0	cc1a9b6ca902
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	67	5	5	5	0	2025-08-28	2025-08-28	5	0	168	168	No	Si	No	Si	Si	No	0	aad92d8af55a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	69	5	5	5	0	2025-08-22	2025-08-22	5	0	167	167	No	Si	No	Si	Si	No	0	37630a1f5591
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	70	4	4	4	0	2026-05-03	2026-07-29	1	3	116	116	No	Si	No	Si	Si	No	0	2310e69de86b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	71	4	4	4	0	2026-07-26	2026-07-26	1	3	235	235	No	Si	No	Si	Si	No	0	ed596bd4737d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	75	4	4	4	0	2026-04-06	2026-04-06	0	4	57	57	No	No	No	Si	No	No	0	7db9ab68ad3e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	77	4	4	4	0	2025-08-20	2025-12-27	0	4	244	244	No	No	No	Si	Si	No	0	7465cdc71c2e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	80	4	4	4	0	2025-08-21	2025-08-24	0	4	288	288	No	No	No	Si	Si	No	0	181475a5cfa0
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	81	3	3	3	0	2026-07-22	2026-07-22	1	2	241	241	No	Si	No	Si	No	No	0	73f4066942ef
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	86	3	3	3	0	2026-02-28	2026-03-05	0	3	237	237	No	Si	No	No	No	No	0	35894a04b485
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	88	3	3	3	0	2026-02-23	2026-02-24	3	0	157	157	No	Si	No	Si	No	No	0	5f12d7222bb2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	89	3	3	3	0	2026-01-25	2026-01-25	0	3	55	55	No	No	No	Si	No	No	0	fa83c9a1a385
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	96	3	3	3	0	2025-11-03	2025-11-11	0	3	184	184	No	Si	No	Si	Si	No	0	ec720b819f02
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	97	3	3	3	0	2025-10-15	2025-10-23	1	2	146	146	No	Si	No	No	No	No	0	eebb3416dbc4
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	98	3	3	3	0	2025-10-03	2025-10-04	3	0	139	139	No	Si	No	Si	Si	No	0	c56db039039a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	99	3	3	3	0	2025-09-26	2025-10-04	0	3	146	146	No	Si	No	No	No	No	0	ef882d661b2a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	100	3	3	3	0	2025-09-29	2025-09-29	3	0	154	154	No	Si	No	Si	Si	No	0	2598cb37080b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	101	3	3	3	0	2025-08-22	2025-08-23	3	0	167	167	No	Si	No	Si	Si	No	0	3bd432a6d6f9
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	102	2	2	2	0	2026-07-29	2026-07-30	0	2	225	225	No	Si	No	Si	No	No	0	728a40759971
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	103	2	2	2	0	2026-07-29	2026-07-29	0	2	239	239	No	Si	No	Si	Si	No	0	5102eaec08df
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	104	2	2	2	0	2026-07-28	2026-07-28	1	1	236	236	No	Si	No	Si	Si	No	0	cd9cb64fe356
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	105	2	2	2	0	2026-07-27	2026-07-28	2	0	225	225	No	Si	No	Si	Si	No	0	7e43fead9a7f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	106	2	2	2	0	2026-07-24	2026-07-24	0	2	226	226	No	Si	No	Si	No	No	0	cb4d29703737
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	107	2	2	2	0	2026-07-21	2026-07-21	0	2	251	251	No	Si	No	Si	Si	No	0	96b8d93330e1
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	108	2	2	2	0	2026-07-19	2026-07-19	1	1	243	243	No	Si	No	Si	No	No	0	cc7950288337
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	109	2	2	2	0	2026-07-18	2026-07-18	0	2	233	233	No	Si	No	Si	No	No	0	de660a962236
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	110	2	2	2	0	2026-07-09	2026-07-16	2	0	117	117	No	No	No	Si	No	No	0	6d0119a649e5
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	111	2	2	2	0	2026-06-24	2026-06-24	0	2	130	130	No	No	No	Si	No	No	0	4dd636733f0b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	112	2	2	2	0	2026-06-24	2026-06-24	0	2	176	176	No	No	No	Si	No	No	0	f0a246a23e62
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	113	2	2	2	0	2026-06-24	2026-06-24	0	2	130	130	No	No	No	Si	No	No	0	2c70930d8f14
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	114	2	2	2	0	2026-06-24	2026-06-24	0	2	176	176	No	No	No	Si	No	No	0	c7c8d20bac09
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	116	2	2	2	0	2026-06-23	2026-06-23	0	2	176	176	No	No	No	Si	No	No	0	2a9e9121f422
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	119	2	2	2	0	2026-05-13	2026-05-14	0	2	160	160	No	Si	No	No	Si	No	0	a85da79df35f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	120	2	2	2	0	2026-05-01	2026-05-10	2	0	252	252	No	No	Si	No	No	No	0	55f27b6749d3
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	121	2	2	2	0	2026-05-03	2026-05-03	0	2	168	168	No	Si	No	No	No	No	0	538bd35c25d2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	123	2	2	2	0	2026-04-11	2026-04-14	2	0	1180	1180	No	Si	No	No	Si	Si	0	e569c8d0ef5f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	124	2	2	2	0	2026-04-12	2026-04-12	0	2	49	49	No	No	No	Si	Si	No	0	d64b7c6a79e9
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	125	2	2	2	0	2026-03-27	2026-03-29	2	0	102	102	No	No	No	Si	No	No	0	4d7d1c3bc638
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	126	2	2	1	1	2026-03-27	2026-03-27	0	2	205	205	No	No	No	Si	No	No	0	d6e2ea467653
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	127	2	2	2	0	2026-03-27	2026-03-27	2	0	256	256	No	No	Si	No	No	No	0	339c8dada2d2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	129	2	2	2	0	2026-02-27	2026-03-04	2	0	204	204	No	No	Si	No	No	No	0	270bea895470
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	132	2	2	2	0	2026-02-03	2026-02-04	0	2	201	201	No	Si	No	No	Si	No	0	216539e6c2e3
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	134	2	2	2	0	2026-01-19	2026-01-21	2	0	226	226	No	No	Si	No	No	No	0	8b97a9239d09
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	135	2	2	2	0	2026-01-12	2026-01-13	2	0	140	140	No	Si	No	No	No	No	0	93f00940bfdf
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	136	2	2	2	0	2025-10-12	2026-01-12	2	0	240	240	No	No	Si	Si	Si	No	0	5a5eb27056f4
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	138	2	2	1	1	2025-12-26	2025-12-29	0	2	163	163	No	Si	No	No	No	No	0	8a8ddd7e6ca0
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	141	2	2	2	0	2025-12-04	2025-12-05	2	0	56	56	No	Si	No	Si	No	No	0	84d8b01a0075
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	142	2	2	2	0	2025-11-29	2025-11-30	2	0	115	115	No	Si	No	No	Si	No	0	a464157fea38
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	144	2	2	2	0	2025-11-22	2025-11-26	2	0	180	180	No	Si	No	No	No	No	0	dbe439f9f158
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	146	2	2	2	0	2025-11-03	2025-11-12	0	2	185	185	No	Si	No	Si	Si	No	0	00f9a7bfd6fe
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	148	2	2	2	0	2025-11-03	2025-11-07	0	2	140	140	No	Si	No	No	No	No	0	1da8cab7a0ac
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	149	2	2	2	0	2025-11-01	2025-11-01	2	0	184	184	No	Si	No	Si	Si	No	0	362bc4041ba4
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	150	2	2	2	0	2025-10-29	2025-10-29	2	0	186	186	No	Si	No	Si	Si	No	0	c730bb80f75f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	151	2	2	2	0	2025-10-24	2025-10-26	0	2	189	189	No	Si	No	Si	Si	No	0	bbe51ba86886
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	152	2	2	2	0	2025-08-24	2025-10-09	0	2	184	184	No	Si	No	Si	Si	No	0	926b7d2f31e9
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	153	2	2	2	0	2025-10-01	2025-10-02	2	0	233	233	No	No	Si	No	No	No	0	b5c91c14f22b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	154	2	2	2	0	2025-10-01	2025-10-01	2	0	163	163	No	Si	No	Si	Si	No	0	24a174716870
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	155	2	2	2	0	2025-09-29	2025-09-29	2	0	58	58	No	Si	No	Si	No	No	0	805266fd1e6f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	156	2	2	2	0	2025-08-28	2025-09-05	0	2	145	145	No	Si	No	No	No	No	0	8d40fd6b8ecf
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	157	2	2	2	0	2025-08-20	2025-08-31	0	2	217	217	No	Si	No	Si	Si	No	0	5b83b680e733
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	159	2	2	2	0	2025-08-25	2025-08-25	2	0	167	167	No	Si	No	Si	Si	No	0	89f59df7f584
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	161	2	2	2	0	2025-08-20	2025-08-21	2	0	145	145	No	Si	No	No	No	No	0	cfe2274485ab
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	164	1	1	1	0	2026-07-29	2026-07-29	1	0	60	60	No	Si	No	Si	No	No	0	a35d95c04836
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	165	1	1	1	0	2026-07-29	2026-07-29	0	1	109	109	No	No	No	Si	Si	No	0	c5189d114e15
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	166	1	1	1	0	2026-07-29	2026-07-29	0	1	109	109	No	No	No	Si	Si	No	0	688039e2e409
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	171	1	1	1	0	2026-07-28	2026-07-28	1	0	235	235	No	Si	No	Si	Si	No	0	4872e7ce4475
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	172	1	1	1	0	2026-07-27	2026-07-27	1	0	116	116	No	No	No	Si	Si	No	0	fa485b66fd8a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	174	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	c8ae533861cb
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	175	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	909d38107023
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	176	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	4a4862df4f95
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	177	1	1	1	0	2026-07-27	2026-07-27	0	1	117	117	No	No	No	Si	Si	No	0	9e288d74fe49
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	178	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	9abd0d32e85c
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	179	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	95f68ff4d982
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	180	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	7be29859013c
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	181	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	b37293beabfa
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	182	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	1d5fd2ba63ec
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	183	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	8b076c0ffbf1
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	184	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	62d5688748a5
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	185	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	95c30f9d2ddd
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	186	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	aaff42647394
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	187	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	ab13861b48dd
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	188	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	3737bb9db87c
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	189	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	376f7929feda
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	190	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	5a98022bce9f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	191	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	141fc468866a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	192	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	ac5cdf0eb1e1
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	193	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	84948accd6d5
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	194	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	35dc5c1ac6b6
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	195	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	7db2702ad7de
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	196	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	0be8c3b14b60
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	197	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	4c30f8fca5f5
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	198	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	7af30ac3423a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	199	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	f6390b458b4f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	200	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	89aa4ac1dc14
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	201	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	4a46fb58ec46
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	202	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	c4734c0d395f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	203	1	1	1	0	2026-07-27	2026-07-27	0	1	109	109	No	No	No	Si	Si	No	0	c999767fd05d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	205	1	1	1	0	2026-07-27	2026-07-27	0	1	117	117	No	No	No	Si	Si	No	0	72e00a66003a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	206	1	1	1	0	2026-07-27	2026-07-27	0	1	109	109	No	No	No	Si	Si	No	0	9bdc2751211a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	207	1	1	1	0	2026-07-27	2026-07-27	0	1	119	119	No	No	No	Si	Si	No	0	344e21d5e827
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	208	1	1	1	0	2026-07-27	2026-07-27	0	1	117	117	No	No	No	Si	Si	No	0	ab30bc37347a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	209	1	1	1	0	2026-07-27	2026-07-27	0	1	109	109	No	No	No	Si	Si	No	0	ae8d2eaa9f3a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	210	1	1	1	0	2026-07-27	2026-07-27	0	1	119	119	No	No	No	Si	Si	No	0	cfb047309b93
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	211	1	1	1	0	2026-07-27	2026-07-27	0	1	117	117	No	No	No	Si	Si	No	0	150f194f65f3
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	212	1	1	1	0	2026-07-27	2026-07-27	0	1	117	117	No	No	No	Si	Si	No	0	997d44370d09
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	213	1	1	1	0	2026-07-27	2026-07-27	0	1	119	119	No	No	No	Si	Si	No	0	b8b3a672c8f3
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	214	1	1	1	0	2026-07-27	2026-07-27	0	1	119	119	No	No	No	Si	Si	No	0	28063f58a67e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	215	1	1	1	0	2026-07-27	2026-07-27	0	1	117	117	No	No	No	Si	Si	No	0	b393fecc26ed
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	216	1	1	1	0	2026-07-27	2026-07-27	0	1	117	117	No	No	No	Si	Si	No	0	c9600d385957
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	217	1	1	1	0	2026-07-27	2026-07-27	0	1	117	117	No	No	No	Si	Si	No	0	69870bf53a83
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	218	1	1	1	0	2026-07-27	2026-07-27	0	1	117	117	No	No	No	Si	Si	No	0	19e39e46e3c5
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	219	1	1	1	0	2026-07-27	2026-07-27	1	0	117	117	No	No	No	Si	Si	No	0	1b6ce74d7cf5
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	220	1	1	1	0	2026-07-27	2026-07-27	0	1	117	117	No	No	No	Si	Si	No	0	64f008b62581
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	221	1	1	1	0	2026-07-27	2026-07-27	0	1	119	119	No	No	No	Si	Si	No	0	30bf3d97caf1
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	222	1	1	1	0	2026-07-27	2026-07-27	0	1	117	117	No	No	No	Si	Si	No	0	0f571db44184
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	223	1	1	1	0	2026-07-27	2026-07-27	0	1	119	119	No	No	No	Si	Si	No	0	23679c083e8f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	224	1	1	1	0	2026-07-27	2026-07-27	0	1	119	119	No	No	No	Si	Si	No	0	7b438e785b45
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	225	1	1	1	0	2026-07-27	2026-07-27	0	1	117	117	No	No	No	Si	Si	No	0	2688fcb50cca
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	226	1	1	1	0	2026-07-27	2026-07-27	0	1	117	117	No	No	No	Si	Si	No	0	a99399c0c51e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	227	1	1	1	0	2026-07-27	2026-07-27	0	1	117	117	No	No	No	Si	Si	No	0	26b3cf2d5847
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	228	1	1	1	0	2026-07-27	2026-07-27	0	1	117	117	No	No	No	Si	Si	No	0	256de8ba9779
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	229	1	1	1	0	2026-07-27	2026-07-27	0	1	116	116	No	No	No	Si	Si	No	0	0f20ac8db6f2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	230	1	1	1	0	2026-07-27	2026-07-27	0	1	119	119	No	No	No	Si	Si	No	0	5268e4dd85a7
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	231	1	1	1	0	2026-07-26	2026-07-26	0	1	94	94	No	Si	No	Si	No	No	0	564bfd9cb678
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	233	1	1	1	0	2026-07-26	2026-07-26	1	0	121	121	No	Si	No	Si	No	No	0	f3b3b13d359b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	234	1	1	1	0	2026-07-25	2026-07-25	0	1	247	247	No	Si	No	Si	Si	No	0	21de5275ceff
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	238	1	1	1	0	2026-07-23	2026-07-23	1	0	229	229	No	Si	No	Si	Si	No	0	649ab41d5bca
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	239	1	1	1	0	2026-07-23	2026-07-23	1	0	60	60	No	No	No	Si	No	No	0	6dce5d3ca71f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	240	1	1	1	0	2026-07-23	2026-07-23	0	1	225	225	No	Si	No	Si	Si	No	0	59dfe40681bc
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	245	1	1	1	0	2026-07-22	2026-07-22	1	0	232	232	No	Si	No	Si	No	No	0	002603446e46
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	247	1	1	1	0	2026-07-21	2026-07-21	0	1	122	122	No	No	No	Si	Si	No	0	67d997a639a5
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	248	1	1	1	0	2026-07-21	2026-07-21	0	1	224	224	No	Si	No	Si	No	No	0	70f27e62c491
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	249	1	1	1	0	2026-07-21	2026-07-21	0	1	221	221	No	Si	No	Si	No	No	0	0baf8fd9cabb
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	250	1	1	1	0	2026-07-21	2026-07-21	0	1	118	118	No	No	No	Si	Si	No	0	fd3dbdbb3a4d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	251	1	1	1	0	2026-07-19	2026-07-19	1	0	244	244	No	Si	No	Si	Si	No	0	b2d4f5a1c5cb
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	253	1	1	1	0	2026-07-19	2026-07-19	0	1	127	127	No	No	No	Si	Si	No	0	7e5c55fc6df4
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	256	1	1	1	0	2026-07-17	2026-07-17	1	0	256	256	No	Si	No	Si	Si	No	0	e6da26598954
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	257	1	1	1	0	2026-07-17	2026-07-17	0	1	127	127	No	No	No	Si	Si	No	0	c9418b0e760c
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	258	1	1	1	0	2026-07-16	2026-07-16	1	0	239	239	No	No	Si	No	Si	No	0	0c7d7c599574
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	259	1	1	1	0	2026-07-16	2026-07-16	1	0	226	226	No	Si	No	Si	No	No	0	2c4592dad407
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	260	1	1	1	0	2026-07-15	2026-07-15	0	1	194	194	No	Si	No	No	Si	No	0	a5e9d83e4213
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	261	1	1	1	0	2026-07-15	2026-07-15	0	1	194	194	No	Si	No	No	Si	No	0	41178520dff4
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	262	1	1	1	0	2026-07-15	2026-07-15	0	1	194	194	No	Si	No	No	Si	No	0	c4f3c421aaa2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	263	1	1	1	0	2026-07-15	2026-07-15	1	0	194	194	No	Si	No	No	Si	No	0	00f8bbebfcdc
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	264	1	1	1	0	2026-07-15	2026-07-15	0	1	194	194	No	Si	No	No	Si	No	0	3951895a262d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	265	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	3df6e11acd52
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	266	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	c3937f01ba01
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	267	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	99d179915b66
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	268	1	1	1	0	2026-07-09	2026-07-09	1	0	194	194	No	Si	No	No	Si	No	0	09f722f745b8
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	269	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	41fd18d6777b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	270	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	5fb419f0a0af
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	271	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	dee335b03e3a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	272	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	b4ed6019be7c
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	273	1	1	1	0	2026-07-09	2026-07-09	1	0	194	194	No	Si	No	No	Si	No	0	6c293acd503a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	274	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	cb50ce1fd50e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	275	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	2a92b0400fa4
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	276	1	1	1	0	2026-07-09	2026-07-09	1	0	194	194	No	Si	No	No	Si	No	0	650075fb0ba5
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	277	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	90ab954df190
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	278	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	92ca5f5e37cc
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	279	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	c0e5c9a70d8d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	280	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	01b689b5619f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	281	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	8f6dc3927515
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	282	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	872702c739cb
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	283	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	06db023eb47c
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	284	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	d57f37fde1da
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	285	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	1905623f3ca4
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	286	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	162a6f78cbd4
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	287	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	a690b0ec2657
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	288	1	1	1	0	2026-07-09	2026-07-09	1	0	194	194	No	Si	No	No	Si	No	0	8b6efa7280fd
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	289	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	b96970bb0b3f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	290	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	4cdcba28e3ef
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	291	1	1	1	0	2026-07-09	2026-07-09	1	0	194	194	No	Si	No	No	Si	No	0	6d18600f4ad8
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	292	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	2b32a888c73a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	293	1	1	1	0	2026-07-09	2026-07-09	0	1	194	194	No	Si	No	No	Si	No	0	8dca63c1d9ea
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	294	1	1	1	0	2026-07-06	2026-07-06	0	1	194	194	No	Si	No	No	Si	No	0	2b0c0c05be47
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	295	1	1	1	0	2026-07-05	2026-07-05	1	0	194	194	No	Si	No	No	Si	No	0	7f8a30fbc561
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	296	1	1	1	0	2026-07-05	2026-07-05	0	1	194	194	No	Si	No	No	Si	No	0	835fc292d7c6
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	297	1	1	1	0	2026-07-04	2026-07-04	1	0	213	213	No	Si	No	No	Si	No	0	e7636a123130
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	298	1	1	1	0	2026-07-03	2026-07-03	0	1	258	258	No	Si	No	No	Si	No	0	c5303c127a43
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	299	1	1	1	0	2026-07-03	2026-07-03	1	0	258	258	No	Si	No	No	Si	No	0	da143a0b5239
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	300	1	1	1	0	2026-07-02	2026-07-02	1	0	258	258	No	Si	No	No	Si	No	0	02b5b95db821
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	301	1	1	1	0	2026-07-02	2026-07-02	1	0	258	258	No	Si	No	No	Si	No	0	92d0826bb4f6
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	302	1	1	1	0	2026-07-02	2026-07-02	0	1	258	258	No	Si	No	No	Si	No	0	af644b01aec9
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	303	1	1	1	0	2026-07-02	2026-07-02	0	1	258	258	No	Si	No	No	Si	No	0	ba385512b3f1
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	304	1	1	1	0	2026-07-02	2026-07-02	1	0	258	258	No	Si	No	No	Si	No	0	105aa10f0ac4
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	305	1	1	1	0	2026-07-02	2026-07-02	1	0	258	258	No	Si	No	No	Si	No	0	dbee3b153527
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	306	1	1	1	0	2026-07-02	2026-07-02	0	1	258	258	No	Si	No	No	Si	No	0	ebd086e90313
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	307	1	1	1	0	2026-07-02	2026-07-02	0	1	130	130	No	No	No	Si	Si	No	0	55713f2479d5
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	308	1	1	1	0	2026-07-02	2026-07-02	0	1	258	258	No	Si	No	No	Si	No	0	47485ebaf20e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	309	1	1	1	0	2026-07-01	2026-07-01	0	1	258	258	No	Si	No	No	Si	No	0	b68ec774dd63
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	310	1	1	1	0	2026-07-01	2026-07-01	0	1	258	258	No	Si	No	No	Si	No	0	60b0a69a790a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	311	1	1	1	0	2026-07-01	2026-07-01	1	0	258	258	No	Si	No	No	Si	No	0	8b2623e756ce
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	312	1	1	1	0	2026-07-01	2026-07-01	1	0	258	258	No	Si	No	No	Si	No	0	3eb30b095692
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	313	1	1	1	0	2026-07-01	2026-07-01	1	0	258	258	No	Si	No	No	Si	No	0	2165a0f12ce2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	314	1	1	1	0	2026-07-01	2026-07-01	0	1	258	258	No	Si	No	No	Si	No	0	7c15c697f1b2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	315	1	1	1	0	2026-07-01	2026-07-01	1	0	258	258	No	Si	No	No	Si	No	0	dbd190e7bb63
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	316	1	1	1	0	2026-07-01	2026-07-01	1	0	258	258	No	Si	No	No	Si	No	0	b46249cb4aeb
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	317	1	1	1	0	2026-07-01	2026-07-01	1	0	258	258	No	Si	No	No	Si	No	0	278e244b7a19
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	318	1	1	1	0	2026-07-01	2026-07-01	0	1	258	258	No	Si	No	No	Si	No	0	329310bd269e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	320	1	1	1	0	2026-07-01	2026-07-01	1	0	258	258	No	Si	No	No	Si	No	0	402d00c4daaa
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	321	1	1	1	0	2026-07-01	2026-07-01	1	0	258	258	No	Si	No	No	Si	No	0	0999f8ed8d98
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	322	1	1	1	0	2026-07-01	2026-07-01	0	1	258	258	No	Si	No	No	Si	No	0	0d056fd46655
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	323	1	1	1	0	2026-07-01	2026-07-01	1	0	258	258	No	Si	No	No	Si	No	0	2cf11346da41
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	324	1	1	1	0	2026-07-01	2026-07-01	0	1	258	258	No	Si	No	No	Si	No	0	5b789b874997
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	325	1	1	1	0	2026-06-30	2026-06-30	0	1	235	235	No	Si	No	No	Si	No	0	118096edd6b7
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	326	1	1	1	0	2026-06-30	2026-06-30	1	0	235	235	No	Si	No	No	Si	No	0	de054dea4e17
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	327	1	1	1	0	2026-06-30	2026-06-30	1	0	235	235	No	Si	No	No	Si	No	0	38cba019c1ec
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	328	1	1	1	0	2026-06-30	2026-06-30	0	1	235	235	No	Si	No	No	Si	No	0	9fea6bf82d0c
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	329	1	1	1	0	2026-06-30	2026-06-30	0	1	235	235	No	Si	No	No	Si	No	0	c16d5db8c43b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	330	1	1	1	0	2026-06-30	2026-06-30	1	0	235	235	No	Si	No	No	Si	No	0	e7c3cd4b0b25
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	331	1	1	1	0	2026-06-30	2026-06-30	0	1	235	235	No	Si	No	No	Si	No	0	140807ea52fb
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	332	1	1	1	0	2026-06-30	2026-06-30	1	0	235	235	No	Si	No	No	Si	No	0	ca9db3175fbc
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	333	1	1	1	0	2026-06-30	2026-06-30	1	0	235	235	No	Si	No	No	Si	No	0	4e8b1df752d6
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	334	1	1	1	0	2026-06-30	2026-06-30	0	1	235	235	No	Si	No	No	Si	No	0	69ee3ca958a7
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	335	1	1	1	0	2026-06-30	2026-06-30	1	0	235	235	No	Si	No	No	Si	No	0	35732253ef45
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	336	1	1	1	0	2026-06-30	2026-06-30	0	1	235	235	No	Si	No	No	Si	No	0	1080aafa09da
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	337	1	1	1	0	2026-06-30	2026-06-30	0	1	235	235	No	Si	No	No	Si	No	0	4910885c0a49
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	338	1	1	1	0	2026-06-30	2026-06-30	1	0	235	235	No	Si	No	No	Si	No	0	d9c0b5f6f945
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	339	1	1	1	0	2026-06-30	2026-06-30	1	0	235	235	No	Si	No	No	Si	No	0	4d8bf94ce172
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	341	1	1	1	0	2026-06-30	2026-06-30	1	0	235	235	No	Si	No	No	Si	No	0	3d026c854920
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	342	1	1	1	0	2026-06-29	2026-06-29	1	0	235	235	No	Si	No	No	Si	No	0	931c0e8514d3
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	343	1	1	1	0	2026-06-29	2026-06-29	0	1	235	235	No	Si	No	No	Si	No	0	4a4bdc870b0f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	344	1	1	1	0	2026-06-29	2026-06-29	1	0	235	235	No	Si	No	No	Si	No	0	fe827a3aea0c
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	345	1	1	1	0	2026-06-29	2026-06-29	0	1	130	130	No	No	No	Si	Si	No	0	a1a2c1c34f68
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	346	1	1	1	0	2026-06-29	2026-06-29	1	0	235	235	No	Si	No	No	Si	No	0	4f73bb87da5e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	347	1	1	1	0	2026-06-28	2026-06-28	0	1	235	235	No	Si	No	No	Si	No	0	02c83fd21b94
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	348	1	1	1	0	2026-06-28	2026-06-28	1	0	235	235	No	Si	No	No	Si	No	0	3b2dc644a96d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	349	1	1	1	0	2026-06-28	2026-06-28	1	0	235	235	No	Si	No	No	Si	No	0	b1b505b0dc83
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	350	1	1	1	0	2026-06-28	2026-06-28	1	0	159	159	No	No	Si	No	No	No	0	374146177ea6
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	351	1	1	1	0	2026-06-28	2026-06-28	1	0	235	235	No	Si	No	No	Si	No	0	b0445eeb09ec
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	352	1	1	1	0	2026-06-28	2026-06-28	0	1	235	235	No	Si	No	No	Si	No	0	8fa70ccaad03
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	353	1	1	1	0	2026-06-28	2026-06-28	0	1	235	235	No	Si	No	No	Si	No	0	3adec176739e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	354	1	1	1	0	2026-06-28	2026-06-28	0	1	235	235	No	Si	No	No	Si	No	0	062a5c47006e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	355	1	1	1	0	2026-06-27	2026-06-27	1	0	235	235	No	Si	No	No	Si	No	0	c5e385496735
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	356	1	1	1	0	2026-06-27	2026-06-27	1	0	235	235	No	Si	No	No	Si	No	0	b28ed5ea6ef1
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	357	1	1	1	0	2026-06-27	2026-06-27	1	0	235	235	No	Si	No	No	Si	No	0	fcac76e9c04e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	358	1	1	1	0	2026-06-27	2026-06-27	1	0	235	235	No	Si	No	No	Si	No	0	c4f6d6626ac3
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	359	1	1	1	0	2026-06-27	2026-06-27	1	0	235	235	No	Si	No	No	Si	No	0	8c19a9379414
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	361	1	1	1	0	2026-06-26	2026-06-26	1	0	235	235	No	Si	No	No	Si	No	0	f9f38313c999
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	362	1	1	1	0	2026-06-26	2026-06-26	0	1	133	133	No	Si	No	Si	No	No	0	a9d4cab6a101
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	364	1	1	1	0	2026-06-26	2026-06-26	0	1	39	39	No	No	No	No	Si	No	0	8fcb27349334
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	365	1	1	1	0	2026-06-26	2026-06-26	0	1	235	235	No	Si	No	No	Si	No	0	943a2a10002b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	366	1	1	1	0	2026-06-26	2026-06-26	1	0	235	235	No	Si	No	No	Si	No	0	58534f23ecce
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	367	1	1	1	0	2026-06-25	2026-06-25	1	0	235	235	No	Si	No	No	Si	No	0	63066ee3e14d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	368	1	1	1	0	2026-06-25	2026-06-25	0	1	235	235	No	Si	No	No	Si	No	0	30e6dd88edd5
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	369	1	1	1	0	2026-06-25	2026-06-25	1	0	235	235	No	Si	No	No	Si	No	0	452a0aba235e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	370	1	1	1	0	2026-06-25	2026-06-25	1	0	235	235	No	Si	No	No	Si	No	0	18070492979d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	371	1	1	1	0	2026-06-25	2026-06-25	0	1	235	235	No	Si	No	No	Si	No	0	5008a44e8244
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	372	1	1	1	0	2026-06-25	2026-06-25	0	1	235	235	No	Si	No	No	Si	No	0	56b4e45dc235
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	373	1	1	1	0	2026-06-25	2026-06-25	0	1	235	235	No	Si	No	No	Si	No	0	f022a54e3563
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	374	1	1	1	0	2026-06-25	2026-06-25	1	0	235	235	No	Si	No	No	Si	No	0	4d6ddb15f21e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	375	1	1	1	0	2026-06-25	2026-06-25	0	1	235	235	No	Si	No	No	Si	No	0	67e80d38af16
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	376	1	1	1	0	2026-06-25	2026-06-25	1	0	235	235	No	Si	No	No	Si	No	0	fdd72902abf5
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	377	1	1	1	0	2026-06-25	2026-06-25	0	1	235	235	No	Si	No	No	Si	No	0	e5a20ac77ebd
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	378	1	1	1	0	2026-06-25	2026-06-25	1	0	235	235	No	Si	No	No	Si	No	0	f7c3d7dae674
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	379	1	1	1	0	2026-06-25	2026-06-25	1	0	235	235	No	Si	No	No	Si	No	0	831d3cc427d3
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	380	1	1	1	0	2026-06-25	2026-06-25	1	0	235	235	No	Si	No	No	Si	No	0	d01a71e0f899
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	381	1	1	1	0	2026-06-25	2026-06-25	1	0	235	235	No	Si	No	No	Si	No	0	7c17c6c2d60e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	382	1	1	1	0	2026-06-25	2026-06-25	1	0	235	235	No	Si	No	No	Si	No	0	ce80e2be9e12
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	383	1	1	1	0	2026-06-25	2026-06-25	0	1	235	235	No	Si	No	No	Si	No	0	2ca0fa4341e2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	384	1	1	1	0	2026-06-25	2026-06-25	0	1	235	235	No	Si	No	No	Si	No	0	25a496d8cb34
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	385	1	1	1	0	2026-06-25	2026-06-25	1	0	235	235	No	Si	No	No	Si	No	0	079cb8cc0ded
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	386	1	1	1	0	2026-06-25	2026-06-25	1	0	235	235	No	Si	No	No	Si	No	0	8157c1dc09df
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	387	1	1	1	0	2026-06-25	2026-06-25	0	1	242	242	No	Si	No	No	Si	No	0	0d5ee653d136
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	392	1	1	1	0	2026-06-24	2026-06-24	1	0	130	130	No	No	No	Si	No	No	0	b82245270cde
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	399	1	1	1	0	2026-06-24	2026-06-24	1	0	130	130	No	No	No	Si	No	No	0	6b1395920330
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	402	1	1	1	0	2026-06-24	2026-06-24	0	1	130	130	No	No	No	Si	No	No	0	df04d26f2221
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	425	1	1	1	0	2026-06-24	2026-06-24	0	1	176	176	No	No	No	Si	No	No	0	962981c3fb7d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	426	1	1	1	0	2026-06-24	2026-06-24	0	1	235	235	No	Si	No	No	Si	No	0	8a13fbdde8f2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	427	1	1	1	0	2026-06-24	2026-06-24	0	1	235	235	No	Si	No	No	Si	No	0	3ccf095a66da
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	430	1	1	1	0	2026-06-23	2026-06-23	0	1	176	176	No	No	No	Si	No	No	0	e14304ec5a28
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	431	1	1	1	0	2026-06-23	2026-06-23	0	1	176	176	No	No	No	Si	No	No	0	7473f01a0559
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	433	1	1	1	0	2026-06-23	2026-06-23	0	1	127	127	No	No	No	Si	No	No	0	3f4d100e3b6b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	437	1	1	1	0	2026-06-22	2026-06-22	0	1	127	127	No	No	No	Si	No	No	0	ef0dc0b51ed9
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	438	1	1	1	0	2026-06-22	2026-06-22	0	1	127	127	No	No	No	Si	No	No	0	d64ea298c363
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	439	1	1	1	0	2026-06-22	2026-06-22	0	1	179	179	No	Si	No	Si	Si	No	0	41cc200ad3d1
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	494	1	1	1	0	2026-06-22	2026-06-22	0	1	168	168	No	No	No	Si	No	No	0	17bc9c6e84ce
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	514	1	1	1	0	2026-06-22	2026-06-22	0	1	127	127	No	No	No	Si	No	No	0	40505824eb03
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	522	1	1	1	0	2026-06-18	2026-06-18	0	1	168	168	No	Si	No	Si	No	No	0	c701dd631615
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	531	1	1	1	0	2026-06-15	2026-06-15	1	0	172	172	No	Si	No	Si	No	No	0	b5a506d47908
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	532	1	1	1	0	2026-06-15	2026-06-15	0	1	172	172	No	Si	No	Si	No	No	0	55337624fff1
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	540	1	1	1	0	2026-06-12	2026-06-12	0	1	168	168	No	No	No	Si	No	No	0	8c3f7fb7d46d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	548	1	1	1	0	2026-06-09	2026-06-09	0	1	304	304	No	No	Si	Si	Si	No	0	19cf6fbf838f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	609	1	1	1	0	2026-05-27	2026-05-27	1	0	249	249	No	No	No	No	Si	No	0	3ec34d6e65f2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	614	1	1	1	0	2026-05-25	2026-05-25	0	1	150	150	No	No	No	Si	No	No	0	1ffaac712850
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	635	1	1	1	0	2026-05-22	2026-05-22	1	0	172	172	No	Si	No	Si	Si	No	0	de9d68816515
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	659	1	1	1	0	2026-05-19	2026-05-19	1	0	57	57	No	No	No	Si	No	No	0	e053c91153bf
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	660	1	1	1	0	2026-05-18	2026-05-18	0	1	210	210	No	Si	No	Si	Si	No	0	b8904308845b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	700	1	1	1	0	2026-05-13	2026-05-13	0	1	149	149	No	No	No	Si	No	No	0	76d7cbdf145d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	751	1	1	1	0	2026-05-11	2026-05-11	0	1	142	142	No	Si	No	No	No	No	0	547f232b29dc
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	859	1	1	1	0	2026-05-08	2026-05-08	0	1	148	148	No	No	No	Si	No	No	0	3d160ab00155
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	866	1	1	1	0	2026-05-06	2026-05-06	1	0	57	57	No	No	No	Si	No	No	0	a6b16ded7834
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	870	1	1	1	0	2026-05-06	2026-05-06	0	1	53	53	No	No	No	Si	No	No	0	579e2ed68091
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	895	1	1	1	0	2026-05-03	2026-05-03	0	1	115	115	No	Si	No	Si	Si	No	0	fa016a252783
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	896	1	1	1	0	2026-05-03	2026-05-03	0	1	147	147	No	No	No	Si	No	No	0	633df783c2ab
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1045	1	1	1	0	2026-04-07	2026-04-07	0	1	124	124	No	No	No	Si	No	No	0	e7c8d12a9eb9
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1046	1	1	1	0	2026-04-07	2026-04-07	0	1	124	124	No	No	No	Si	No	No	0	72ff885a6bbb
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1055	1	1	1	0	2026-04-06	2026-04-06	1	0	258	258	No	No	Si	No	No	No	0	3006b0a266bb
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1056	1	1	1	0	2026-04-06	2026-04-06	1	0	212	212	No	No	Si	No	No	No	0	5d5b2ae7420a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1057	1	1	1	0	2026-04-06	2026-04-06	1	0	121	121	No	No	No	No	No	No	0	b62ce56e76b6
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1076	1	1	1	0	2026-04-02	2026-04-02	1	0	153	153	No	No	No	No	No	No	0	ed9f669cc774
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1078	1	1	1	0	2026-04-02	2026-04-02	1	0	421	421	No	No	Si	No	Si	No	0	f1cca533bbc9
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1080	1	1	1	0	2026-04-02	2026-04-02	1	0	433	433	No	No	Si	No	Si	No	0	187ffcfee10b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1081	1	1	1	0	2026-04-02	2026-04-02	1	0	445	445	No	No	Si	No	Si	No	0	eccc69cc69e1
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1082	1	1	1	0	2026-04-02	2026-04-02	1	0	429	429	No	No	Si	No	Si	No	0	06de6d54d5d6
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1086	1	1	1	0	2026-04-01	2026-04-01	0	1	51	51	No	Si	No	Si	No	No	0	6ad81b60156c
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1087	1	1	1	0	2026-03-31	2026-03-31	1	0	102	102	No	No	No	No	No	No	0	10856bc1ca40
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1089	1	1	1	0	2026-03-31	2026-03-31	1	0	161	161	No	No	No	No	No	No	0	f353545e5c79
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1091	1	1	1	0	2026-03-31	2026-03-31	1	0	421	421	No	No	Si	No	Si	No	0	20f79d4beaa9
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1092	1	1	1	0	2026-03-31	2026-03-31	1	0	425	425	No	No	Si	No	Si	No	0	07e85e1ed6de
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1093	1	1	1	0	2026-03-31	2026-03-31	1	0	433	433	No	No	Si	No	Si	No	0	e90f14b75b97
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1094	1	1	1	0	2026-03-31	2026-03-31	1	0	421	421	No	No	Si	No	Si	No	0	2f0d6e340104
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1095	1	1	1	0	2026-03-31	2026-03-31	1	0	441	441	No	No	Si	No	Si	No	0	87e4d36b05ad
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1096	1	1	1	0	2026-03-31	2026-03-31	1	0	429	429	No	No	Si	No	Si	No	0	f61258d2e036
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1097	1	1	1	0	2026-03-31	2026-03-31	1	0	437	437	No	No	Si	No	Si	No	0	a626dbba9c30
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1098	1	1	1	0	2026-03-31	2026-03-31	1	0	425	425	No	No	Si	No	Si	No	0	d313ae47f84e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1099	1	1	1	0	2026-03-31	2026-03-31	1	0	441	441	No	No	Si	No	Si	No	0	60e117e3dbe2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1100	1	1	1	0	2026-03-31	2026-03-31	1	0	425	425	No	No	Si	No	Si	No	0	709f987ce7ec
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1101	1	1	1	0	2026-03-31	2026-03-31	1	0	421	421	No	No	Si	No	Si	No	0	b422f4f1d214
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1102	1	1	1	0	2026-03-31	2026-03-31	1	0	425	425	No	No	Si	No	Si	No	0	abb88a1adca9
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1103	1	1	1	0	2026-03-31	2026-03-31	1	0	421	421	No	No	Si	No	Si	No	0	0e00b62613ae
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1104	1	1	1	0	2026-03-31	2026-03-31	1	0	441	441	No	No	Si	No	Si	No	0	366739b0df4e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1105	1	1	1	0	2026-03-31	2026-03-31	1	0	433	433	No	No	Si	No	Si	No	0	1e4e1edd64ef
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1106	1	1	1	0	2026-03-31	2026-03-31	1	0	441	441	No	No	Si	No	Si	No	0	1c943edb7dfc
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1107	1	1	1	0	2026-03-31	2026-03-31	1	0	100	100	No	No	No	No	No	No	0	212c1f6ca1f8
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1114	1	1	1	0	2026-03-30	2026-03-30	1	0	139	139	No	Si	No	No	Si	No	0	78dbf49cb6e3
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1117	1	1	1	0	2026-03-29	2026-03-29	1	0	110	110	No	No	No	No	No	No	0	eaaeee07a57a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1118	1	1	1	0	2026-03-29	2026-03-29	1	0	53	53	No	No	No	Si	No	No	0	b7a1f8945550
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1119	1	1	1	0	2026-03-28	2026-03-28	1	0	137	137	No	Si	No	No	Si	No	0	7d7c843159a0
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1120	1	1	1	0	2026-03-28	2026-03-28	1	0	138	138	No	Si	No	No	Si	No	0	eb4bfa18e54f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1121	1	1	1	0	2026-03-27	2026-03-27	1	0	886	886	No	No	Si	No	No	No	0	e8301333bbe6
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1126	1	1	1	0	2026-03-26	2026-03-26	1	0	120	120	No	No	No	No	No	No	0	c5e20367be7a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1127	1	1	1	0	2026-03-26	2026-03-26	1	0	84	84	No	No	No	Si	Si	No	0	79c6d1468e91
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1135	1	1	1	0	2026-03-25	2026-03-25	1	0	71	71	No	No	No	No	Si	No	0	a1b406bdff0f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1138	1	1	1	0	2026-03-24	2026-03-24	1	0	239	239	No	No	Si	No	No	No	0	f11f140a0cf6
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1143	1	1	1	0	2026-03-23	2026-03-23	1	0	88	88	No	No	No	Si	Si	No	0	72c45a7b200b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1148	1	1	1	0	2026-03-21	2026-03-21	1	0	260	260	No	No	No	Si	Si	No	0	f08b407a144d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1149	1	1	1	0	2026-03-21	2026-03-21	1	0	199	199	No	No	No	Si	No	No	0	d132ea563037
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1150	1	1	1	0	2026-03-21	2026-03-21	1	0	256	256	No	No	Si	No	No	No	0	8e5097f306a7
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1151	1	1	1	0	2026-03-17	2026-03-17	1	0	206	206	No	Si	No	No	No	No	0	15ad0cd1c4d3
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1154	1	1	1	0	2026-03-16	2026-03-16	0	1	206	206	No	Si	No	No	No	No	0	51dbae2a66be
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1155	1	1	1	0	2026-03-16	2026-03-16	0	1	206	206	No	Si	No	No	No	No	0	83c522b8808d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1156	1	1	1	0	2026-03-16	2026-03-16	0	1	206	206	No	Si	No	No	No	No	0	acdd74a7042d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1157	1	1	1	0	2026-03-16	2026-03-16	0	1	206	206	No	Si	No	No	No	No	0	1d1c123a5fcd
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1158	1	1	1	0	2026-03-16	2026-03-16	0	1	206	206	No	Si	No	No	No	No	0	885acb528a5d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1159	1	1	1	0	2026-03-16	2026-03-16	0	1	206	206	No	Si	No	No	No	No	0	9bc02d77018e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1160	1	1	1	0	2026-03-16	2026-03-16	0	1	206	206	No	Si	No	No	No	No	0	57c560f96775
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1162	1	1	1	0	2026-03-15	2026-03-15	0	1	206	206	No	Si	No	No	No	No	0	8dcde07617f2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1163	1	1	1	0	2026-03-15	2026-03-15	1	0	117	117	No	Si	No	No	No	No	0	f241987273f7
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1164	1	1	1	0	2026-03-15	2026-03-15	0	1	206	206	No	Si	No	No	No	No	0	645093e1228d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1165	1	1	1	0	2026-03-15	2026-03-15	0	1	206	206	No	Si	No	No	No	No	0	384563dbee65
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1166	1	1	1	0	2026-03-15	2026-03-15	1	0	445	445	No	No	Si	No	Si	No	0	d612c8c2ca7b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1167	1	1	1	0	2026-03-15	2026-03-15	1	0	437	437	No	No	Si	No	Si	No	0	85d273da8282
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1168	1	1	1	0	2026-03-15	2026-03-15	1	0	441	441	No	No	Si	No	Si	No	0	ead781b9745c
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1170	1	1	1	0	2026-03-15	2026-03-15	1	0	425	425	No	No	Si	No	Si	No	0	fc5dd6fa4800
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1171	1	1	1	0	2026-03-15	2026-03-15	1	0	433	433	No	No	Si	No	Si	No	0	8833c75ceac2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1172	1	1	1	0	2026-03-15	2026-03-15	1	0	433	433	No	No	Si	No	Si	No	0	399a2ea04e70
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1173	1	1	1	0	2026-03-15	2026-03-15	1	0	425	425	No	No	Si	No	Si	No	0	a6e742a4a5ac
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1174	1	1	1	0	2026-03-15	2026-03-15	1	0	425	425	No	No	Si	No	Si	No	0	b01acf93b153
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1175	1	1	1	0	2026-03-15	2026-03-15	1	0	409	409	No	No	Si	No	Si	No	0	c1cbd1f98a98
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1176	1	1	1	0	2026-03-15	2026-03-15	1	0	425	425	No	No	Si	No	Si	No	0	c5a68143edd3
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1177	1	1	1	0	2026-03-15	2026-03-15	1	0	417	417	No	No	Si	No	Si	No	0	391306abcd23
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1178	1	1	1	0	2026-03-14	2026-03-14	1	0	137	137	No	No	No	Si	No	No	0	b8e5b702f550
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1179	1	1	1	0	2026-03-14	2026-03-14	1	0	117	117	No	Si	No	No	No	No	0	55a8596caae0
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1182	1	1	1	0	2026-03-14	2026-03-14	1	0	116	116	No	Si	No	No	No	No	0	ec1a7a94b269
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1184	1	1	1	0	2026-03-12	2026-03-12	0	1	206	206	No	Si	No	No	No	No	0	01008a716b5f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1186	1	1	1	0	2026-03-12	2026-03-12	0	1	57	57	No	No	No	Si	No	No	0	59d79c89367e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1188	1	1	1	0	2026-03-12	2026-03-12	1	0	149	149	No	Si	No	No	No	No	0	d0e922ea8011
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1190	1	1	1	0	2026-03-12	2026-03-12	0	1	145	145	No	Si	No	Si	No	No	0	45a475401ecb
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1191	1	1	1	0	2026-03-12	2026-03-12	0	1	151	151	No	Si	No	Si	No	No	0	1785bbc16335
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1194	1	1	1	0	2026-03-10	2026-03-10	0	1	57	57	No	No	No	Si	No	No	0	c20e60045825
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1196	1	1	1	0	2026-03-10	2026-03-10	0	1	57	57	No	No	No	Si	No	No	0	2422442b3a75
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1225	1	1	1	0	2026-03-08	2026-03-08	1	0	61	61	No	No	No	Si	No	No	0	8e4723af089a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1227	1	1	1	0	2026-03-08	2026-03-08	0	1	60	60	No	No	No	Si	No	No	0	ea6a9aa8f011
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1228	1	1	1	0	2026-03-08	2026-03-08	0	1	106	106	No	No	No	Si	Si	No	0	b34b565259c2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1229	1	1	1	0	2026-03-06	2026-03-06	0	1	108	108	No	No	No	Si	No	No	0	8cd0dfd751bb
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1231	1	1	1	0	2026-03-05	2026-03-05	1	0	101	101	No	No	No	Si	No	No	0	edc4b29a2e8e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1232	1	1	1	0	2026-03-05	2026-03-05	1	0	100	100	No	No	No	Si	No	No	0	d41a4655b1aa
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1233	1	1	1	0	2026-03-05	2026-03-05	0	1	82	82	No	No	No	Si	No	No	0	df924ba21ab7
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1234	1	1	1	0	2026-03-05	2026-03-05	0	1	234	234	No	Si	No	No	Si	No	0	8241cdbbbe8e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1235	1	1	1	0	2026-03-05	2026-03-05	1	0	210	210	No	No	No	Si	Si	No	0	966373619eff
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1240	1	1	1	0	2026-02-26	2026-02-26	0	1	257	257	No	No	No	No	No	No	0	b094ff85316f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1241	1	1	1	0	2026-02-25	2026-02-25	0	1	704	704	No	Si	No	No	No	No	0	bfca016624ec
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1242	1	1	1	0	2026-02-22	2026-02-22	1	0	54	54	No	No	No	Si	No	No	0	bba2b19a6a77
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1243	1	1	1	0	2026-02-22	2026-02-22	1	0	56	56	No	No	No	Si	No	No	0	cd7096683980
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1246	1	1	0	1	2026-02-21	2026-02-21	0	1	65	65	No	No	No	Si	No	No	0	ee9899fff900
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1247	1	1	1	0	2026-02-18	2026-02-18	1	0	53	53	No	No	No	Si	No	No	0	8be21c65c416
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1249	1	1	0	1	2026-02-17	2026-02-17	0	1	184	184	No	No	No	No	Si	No	0	b4608a776452
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1250	1	1	1	0	2026-02-10	2026-02-10	1	0	60	60	No	No	No	Si	No	No	0	58d07bc1d098
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1252	1	1	1	0	2026-02-02	2026-02-02	1	0	55	55	No	No	No	Si	No	No	0	bcd60bf7a066
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1253	1	1	0	1	2026-01-31	2026-01-31	0	1	54	54	No	No	No	Si	No	No	0	5fcc75aba719
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1254	1	1	1	0	2026-01-27	2026-01-27	1	0	143	143	No	Si	No	No	No	No	0	ec7e0533edec
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1260	1	1	1	0	2026-01-17	2026-01-17	0	1	266	266	No	Si	No	Si	Si	No	0	488ee093ccb9
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1261	1	1	1	0	2026-01-17	2026-01-17	0	1	249	249	No	No	No	Si	No	No	0	17e288ae7521
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1262	1	1	1	0	2026-01-17	2026-01-17	0	1	243	243	No	No	No	Si	No	No	0	5b82480470dc
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1267	1	1	1	0	2026-01-13	2026-01-13	1	0	55	55	No	No	No	Si	No	No	0	a268059891ff
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1268	1	1	1	0	2026-01-12	2026-01-12	1	0	280	280	No	No	Si	No	Si	No	0	5a33e5aca028
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1272	1	1	1	0	2026-01-11	2026-01-11	0	1	229	229	No	Si	No	Si	Si	No	0	c7ceef78b94b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1288	1	1	1	0	2026-01-09	2026-01-09	1	0	107	107	No	No	No	No	No	No	0	11d31ddf114d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1289	1	1	1	0	2026-01-09	2026-01-09	1	0	120	120	No	No	No	No	No	No	0	35200a15ebe5
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1290	1	1	1	0	2026-01-09	2026-01-09	1	0	55	55	No	No	No	Si	No	No	0	d792b0a0ece9
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1291	1	1	1	0	2026-01-09	2026-01-09	0	1	202	202	No	Si	No	No	No	No	0	150910f2dee2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1292	1	1	1	0	2026-01-09	2026-01-09	1	0	55	55	No	No	No	Si	No	No	0	984fd1cae4cc
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1294	1	1	1	0	2026-01-08	2026-01-08	1	0	231	231	No	No	Si	No	No	No	0	2e85dfd28d36
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1295	1	1	1	0	2026-01-05	2026-01-05	1	0	121	121	No	No	No	No	No	No	0	8963a56f39b1
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1298	1	1	1	0	2026-01-04	2026-01-04	0	1	254	254	No	No	No	No	No	No	0	064dff39140d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1299	1	1	1	0	2026-01-04	2026-01-04	1	0	254	254	No	No	No	Si	Si	No	0	65c40c2d6faf
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1302	1	1	1	0	2025-12-30	2025-12-30	0	1	60	60	No	No	No	Si	No	No	0	006994606eff
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1303	1	1	1	0	2025-12-29	2025-12-29	1	0	54	54	No	No	No	Si	No	No	0	dda3fe02254a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1305	1	1	1	0	2025-12-23	2025-12-23	0	1	204	204	No	Si	No	Si	Si	No	0	6c2a0a57c6ff
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1306	1	1	1	0	2025-12-23	2025-12-23	0	1	283	283	No	Si	No	No	No	No	0	1e4b8ff85243
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1308	1	1	1	0	2025-12-17	2025-12-17	0	1	177	177	No	Si	No	Si	Si	No	0	9164a58e5138
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1311	1	1	1	0	2025-12-12	2025-12-12	1	0	154	154	No	Si	No	Si	No	No	0	9cfa71560612
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1315	1	1	1	0	2025-12-09	2025-12-09	1	0	207	207	No	No	Si	No	No	No	0	67aa6d745419
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1321	1	1	1	0	2025-12-05	2025-12-05	1	0	207	207	No	No	Si	No	No	No	0	358df39f8aee
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1322	1	1	1	0	2025-12-05	2025-12-05	0	1	196	196	No	Si	No	Si	Si	No	0	2839cb495e4e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1323	1	1	0	1	2025-12-05	2025-12-05	1	0	141	141	No	Si	No	Si	No	No	0	8d163e53d560
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1324	1	1	1	0	2025-12-05	2025-12-05	1	0	142	142	No	Si	No	Si	No	No	0	002260c3dfd6
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1327	1	1	1	0	2025-12-03	2025-12-03	0	1	220	220	No	Si	No	Si	Si	No	0	8604837c54a2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1328	1	1	1	0	2025-12-02	2025-12-02	0	1	208	208	No	No	No	Si	Si	No	0	4668a5c4d965
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1331	1	1	1	0	2025-11-29	2025-11-29	1	0	229	229	No	No	Si	No	No	No	0	37f5c59d2069
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1332	1	1	1	0	2025-11-26	2025-11-26	0	1	235	235	No	Si	No	No	Si	No	0	f31cc226a710
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1338	1	1	1	0	2025-11-22	2025-11-22	0	1	193	193	No	Si	No	Si	Si	No	0	172065b72a39
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1339	1	1	1	0	2025-11-21	2025-11-21	0	1	194	194	No	Si	No	Si	Si	No	0	6048f920283a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1341	1	1	1	0	2025-11-21	2025-11-21	0	1	141	141	No	Si	No	Si	No	No	0	85a52837ac3b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1342	1	1	1	0	2025-11-21	2025-11-21	1	0	177	177	No	Si	No	No	No	No	0	a552f86c6c6b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1343	1	1	1	0	2025-11-21	2025-11-21	1	0	165	165	No	Si	No	No	No	No	0	0354e70a7a45
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1344	1	1	1	0	2025-11-17	2025-11-17	0	1	199	199	No	Si	No	Si	Si	No	0	77a5d6c6fea0
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1345	1	1	1	0	2025-11-16	2025-11-16	0	1	254	254	No	Si	No	Si	Si	No	0	7505cf98bb2c
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1347	1	1	1	0	2025-11-14	2025-11-14	0	1	205	205	No	Si	No	Si	Si	No	0	3f800451dced
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1348	1	1	1	0	2025-11-14	2025-11-14	1	0	107	107	No	Si	No	No	No	No	0	1c1706798eed
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1353	1	1	1	0	2025-11-04	2025-11-04	0	1	138	138	No	Si	No	Si	Si	No	0	2bfe78367548
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1354	1	1	1	0	2025-11-04	2025-11-04	0	1	141	141	No	Si	No	Si	Si	No	0	60088245496c
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1355	1	1	1	0	2025-11-04	2025-11-04	0	1	213	213	No	Si	No	Si	Si	No	0	8964264f3ba6
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1360	1	1	1	0	2025-10-30	2025-10-30	0	1	174	174	No	Si	Si	Si	Si	No	0	ebca8bcf839a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1364	1	1	1	0	2025-10-27	2025-10-27	1	0	229	229	No	No	Si	No	No	No	0	0f30c66b049a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1366	1	1	1	0	2025-10-22	2025-10-22	0	1	178	178	No	Si	No	Si	Si	No	0	f7536e5774c3
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1369	1	1	1	0	2025-10-22	2025-10-22	0	1	174	174	No	Si	No	Si	Si	No	0	7f6938d2f076
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1373	1	1	1	0	2025-10-14	2025-10-14	0	1	204	204	No	Si	No	Si	Si	No	0	1cf6fda857c0
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1375	1	1	1	0	2025-10-11	2025-10-11	0	1	227	227	No	Si	No	No	Si	No	0	b10a879701c1
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1377	1	1	1	0	2025-10-09	2025-10-09	0	1	206	206	No	Si	No	No	Si	No	0	2b1f8d3467c0
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1378	1	1	1	0	2025-10-09	2025-10-09	1	0	233	233	No	No	Si	No	No	No	0	af3d3387b96f
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1381	1	1	1	0	2025-10-08	2025-10-08	0	1	198	198	No	Si	No	No	Si	No	0	1ee004c32ded
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1382	1	1	1	0	2025-10-08	2025-10-08	0	1	481	481	No	Si	No	No	Si	No	0	37b13a4e72d2
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1383	1	1	1	0	2025-10-05	2025-10-05	1	0	175	175	No	No	Si	Si	No	No	0	54b01e8837d1
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1384	1	1	1	0	2025-10-05	2025-10-05	0	1	250	250	No	Si	Si	Si	Si	No	0	a93b285abe55
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1385	1	1	1	0	2025-10-04	2025-10-04	1	0	141	141	No	Si	No	Si	Si	No	0	713202c904d0
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1386	1	1	1	0	2025-10-03	2025-10-03	0	1	148	148	No	Si	No	No	Si	No	0	e35d02433c84
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1387	1	1	1	0	2025-10-02	2025-10-02	1	0	209	209	No	No	Si	No	No	No	0	88c62019aefa
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1388	1	1	1	0	2025-10-02	2025-10-02	1	0	150	150	No	Si	No	Si	Si	No	0	6edbc3f4d905
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1391	1	1	1	0	2025-10-02	2025-10-02	1	0	217	217	No	No	Si	No	No	No	0	400f84c5031c
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1394	1	1	1	0	2025-10-01	2025-10-01	1	0	237	237	No	No	Si	No	No	No	0	def668e21ff5
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1395	1	1	1	0	2025-09-30	2025-09-30	1	0	217	217	No	No	Si	No	No	No	0	fe5f8caf8b39
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1396	1	1	1	0	2025-09-28	2025-09-28	1	0	155	155	No	Si	No	Si	Si	No	0	08bff557b856
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1397	1	1	1	0	2025-09-27	2025-09-27	1	0	237	237	No	No	Si	No	No	No	0	a882d333188a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1398	1	1	1	0	2025-09-23	2025-09-23	1	0	163	163	No	No	No	Si	No	No	0	c4d214067484
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1399	1	1	1	0	2025-09-19	2025-09-19	1	0	197	197	No	Si	No	Si	No	No	0	4317b0d369a3
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1400	1	1	1	0	2025-09-19	2025-09-19	1	0	220	220	No	Si	No	Si	Si	No	0	4c090d0722a7
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1401	1	1	1	0	2025-09-16	2025-09-16	0	1	147	147	No	Si	No	No	No	No	0	e82e44ea5b18
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1403	1	1	1	0	2025-09-14	2025-09-14	1	0	241	241	No	No	Si	No	No	No	0	6fcec68088db
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1404	1	1	1	0	2025-09-12	2025-09-12	1	0	237	237	No	No	Si	No	No	No	0	1490f2979ba3
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1406	1	1	1	0	2025-09-11	2025-09-11	0	1	263	263	No	Si	No	No	No	No	0	cd5527351d84
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1407	1	1	1	0	2025-09-08	2025-09-08	0	1	229	229	No	Si	No	Si	Si	No	0	d6f3ad5c63e5
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1410	1	1	1	0	2025-09-04	2025-09-04	1	0	191	191	No	No	No	No	No	No	0	7dd82f899c40
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1413	1	1	1	0	2025-09-04	2025-09-04	0	1	229	229	No	Si	No	Si	Si	No	0	2d6cbeb8b439
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1415	1	1	1	0	2025-09-03	2025-09-03	0	1	196	196	No	Si	No	Si	No	No	0	a885a9ed67bf
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1418	1	1	1	0	2025-09-02	2025-09-02	1	0	207	207	No	Si	No	Si	Si	No	0	5a2cf8e68c90
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1422	1	1	1	0	2025-09-01	2025-09-01	1	0	201	201	No	Si	No	Si	Si	No	0	9614099e76a5
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1426	1	1	1	0	2025-08-31	2025-08-31	1	0	201	201	No	Si	No	Si	Si	No	0	5a5e6b2d2329
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1430	1	1	1	0	2025-08-29	2025-08-29	0	1	194	194	No	Si	No	Si	Si	No	0	8edac6c37f0b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1431	1	1	1	0	2025-08-29	2025-08-29	1	0	247	247	No	No	No	No	No	No	0	f535e77f8ac3
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1433	1	1	1	0	2025-08-28	2025-08-28	1	0	247	247	No	No	No	No	No	No	0	403571bb994d
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1444	1	1	1	0	2025-08-27	2025-08-27	1	0	168	168	No	Si	No	Si	Si	No	0	5474df5573ff
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1448	1	1	1	0	2025-08-26	2025-08-26	1	0	185	185	No	No	No	No	No	No	0	9ac797b10826
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1455	1	1	1	0	2025-08-25	2025-08-25	0	1	90	90	No	Si	No	Si	No	No	0	01d92a21663e
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1456	1	1	1	0	2025-08-25	2025-08-25	1	0	170	170	No	No	Si	No	No	No	0	7a514660c8c4
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1464	1	1	1	0	2025-08-23	2025-08-23	1	0	182	182	No	No	Si	No	No	No	0	81bc5dea784a
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1465	1	1	1	0	2025-08-22	2025-08-22	1	0	181	181	No	Si	No	Si	Si	No	0	590d01c4c1ff
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1466	1	1	1	0	2025-08-22	2025-08-22	1	0	192	192	No	Si	No	Si	No	No	0	4de5b6f898c9
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1467	1	1	1	0	2025-08-22	2025-08-22	0	1	173	173	No	Si	No	Si	Si	No	0	5eade782faf1
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1468	1	1	1	0	2025-08-22	2025-08-22	0	1	168	168	No	Si	No	Si	Si	No	0	bfe3a1467be6
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1469	1	1	1	0	2025-08-22	2025-08-22	0	1	247	247	No	No	No	Si	Si	No	0	c26224227d5c
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1472	1	1	1	0	2025-08-22	2025-08-22	0	1	229	229	No	Si	No	Si	Si	No	0	7f657b24dd22
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1473	1	1	1	0	2025-08-21	2025-08-21	1	0	177	177	No	Si	No	Si	Si	No	0	399fd322b940
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1474	1	1	1	0	2025-08-21	2025-08-21	1	0	179	179	No	Si	No	Si	Si	No	0	d490001685bd
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1475	1	1	1	0	2025-08-21	2025-08-21	1	0	178	178	No	No	Si	No	No	No	0	10f22a3a7beb
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1476	1	1	1	0	2025-08-20	2025-08-20	0	1	187	187	No	Si	No	Si	Si	No	0	9cdaab340178
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1477	1	1	1	0	2025-08-20	2025-08-20	0	1	245	245	No	No	No	No	No	No	0	456a1f9fe62b
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1478	1	1	1	0	2025-08-20	2025-08-20	1	0	174	174	No	No	Si	No	No	No	0	53e4730786a3
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1479	1	1	1	0	2025-08-20	2025-08-20	0	1	171	171	No	Si	No	Si	Si	No	0	2f37c6c337f5
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1480	1	1	1	0	2025-08-20	2025-08-20	1	0	178	178	No	Si	No	Si	Si	No	0	f0f6e23638af
08 - powershell (linea de comandos)	08	powershell (linea de comandos)	1482	1	1	1	0	2025-08-20	2025-08-20	0	1	215	215	No	Si	No	Si	Si	No	0	2489ca3c65ff
09 - cmd.exe	09	cmd.exe	4	151	151	63	88	2025-12-26	2026-02-28	0	151	132	132	No	Si	No	No	No	No	0	76766431d3a8
09 - cmd.exe	09	cmd.exe	5	137	137	135	1	2025-12-07	2026-02-28	1	136	219	219	No	Si	No	No	Si	No	1	43dac4ab6937
09 - cmd.exe	09	cmd.exe	11	64	64	31	33	2025-12-30	2026-02-28	0	64	116	116	No	Si	No	No	No	No	0	e21ea0cd0f53
09 - cmd.exe	09	cmd.exe	14	44	44	42	2	2025-12-06	2026-02-06	0	44	230	230	No	Si	No	No	Si	No	0	8d6cba304c76
09 - cmd.exe	09	cmd.exe	16	35	35	31	4	2025-12-26	2026-02-28	0	35	178	178	No	Si	No	No	No	No	0	4aebddb06198
09 - cmd.exe	09	cmd.exe	19	31	31	28	3	2025-12-28	2026-02-28	0	31	175	175	No	Si	No	No	No	No	0	1179e3722c16
09 - cmd.exe	09	cmd.exe	32	19	19	19	0	2025-12-13	2026-02-28	0	19	196	196	No	Si	No	No	No	No	0	9246fb058792
09 - cmd.exe	09	cmd.exe	34	18	18	18	0	2025-12-15	2025-12-15	0	18	50	50	No	Si	No	No	No	No	0	b02dab4084cc
09 - cmd.exe	09	cmd.exe	35	18	18	18	0	2025-12-04	2025-12-06	0	18	51	51	No	Si	No	No	No	No	0	becf49c7d5dd
09 - cmd.exe	09	cmd.exe	38	12	12	12	0	2025-12-05	2025-12-18	0	12	249	249	No	Si	No	No	Si	No	0	a755bb5a3829
09 - cmd.exe	09	cmd.exe	43	9	9	9	0	2025-12-17	2026-01-25	0	9	175	175	No	Si	No	No	No	No	0	fe5fdf1f305b
09 - cmd.exe	09	cmd.exe	45	8	8	8	0	2026-06-20	2026-07-16	8	0	238	238	No	Si	No	Si	No	No	0	dbda634bbf3f
09 - cmd.exe	09	cmd.exe	49	7	7	6	1	2025-12-08	2025-12-08	0	7	196	196	No	Si	No	No	No	No	0	5d13b467e29f
09 - cmd.exe	09	cmd.exe	52	6	6	6	0	2026-03-19	2026-03-21	2	4	131	131	No	Si	No	No	No	No	0	0f713a9472e3
09 - cmd.exe	09	cmd.exe	62	5	5	5	0	2026-01-14	2026-02-02	0	5	117	117	No	Si	No	No	No	No	0	d422d0b2ecd1
09 - cmd.exe	09	cmd.exe	76	4	4	4	0	2026-03-29	2026-04-04	0	4	147	147	No	Si	No	No	No	No	0	ef18a4afde7c
09 - cmd.exe	09	cmd.exe	90	3	3	3	0	2026-01-17	2026-01-17	3	0	220	220	No	Si	No	No	No	No	0	87d1ff6a8b34
09 - cmd.exe	09	cmd.exe	92	3	3	3	0	2025-09-19	2026-01-11	0	3	193	193	No	Si	No	No	No	No	0	23b63bfc7cbb
09 - cmd.exe	09	cmd.exe	95	3	3	3	0	2025-11-13	2025-11-14	0	3	48	48	No	Si	No	No	No	No	0	225a8c719f7a
09 - cmd.exe	09	cmd.exe	115	2	2	2	0	2026-06-23	2026-06-23	2	0	144	144	No	Si	No	No	No	No	0	f76a1426e864
09 - cmd.exe	09	cmd.exe	122	2	2	1	1	2026-05-01	2026-05-01	0	2	123	123	No	Si	No	No	No	No	0	a4681ab13551
09 - cmd.exe	09	cmd.exe	139	2	2	2	0	2025-12-08	2025-12-08	0	2	51	51	No	Si	No	No	No	No	0	e7166afadaa8
09 - cmd.exe	09	cmd.exe	140	2	2	2	0	2025-12-06	2025-12-06	0	2	50	50	No	Si	No	No	No	No	0	cfa1bfcc2df1
09 - cmd.exe	09	cmd.exe	143	2	2	2	0	2025-11-27	2025-11-27	0	2	256	256	No	Si	No	No	Si	No	0	60fcb105b1f1
09 - cmd.exe	09	cmd.exe	145	2	2	2	0	2025-08-24	2025-11-23	0	2	163	163	No	Si	No	Si	Si	No	0	ce851ed7de46
09 - cmd.exe	09	cmd.exe	147	2	2	2	0	2025-11-08	2025-11-10	2	0	261	261	No	Si	No	No	No	No	0	30375f7de55c
09 - cmd.exe	09	cmd.exe	242	1	1	1	0	2026-07-22	2026-07-22	0	1	187	187	No	Si	No	No	No	No	0	adc40dbd74b0
09 - cmd.exe	09	cmd.exe	243	1	1	1	0	2026-07-22	2026-07-22	0	1	187	187	No	Si	No	No	No	No	0	44de3f7d8535
09 - cmd.exe	09	cmd.exe	601	1	1	1	0	2026-06-01	2026-06-01	1	0	169	169	No	Si	No	Si	Si	No	0	1933630caf56
09 - cmd.exe	09	cmd.exe	615	1	1	1	0	2026-05-25	2026-05-25	1	0	143	143	No	No	No	No	No	No	0	122cebae05d6
09 - cmd.exe	09	cmd.exe	647	1	1	1	0	2026-05-20	2026-05-20	1	0	245	245	No	No	No	No	Si	No	0	c4dab1177c50
09 - cmd.exe	09	cmd.exe	740	1	1	1	0	2026-05-12	2026-05-12	1	0	181	181	No	Si	No	No	No	No	0	9605a0d4d4be
09 - cmd.exe	09	cmd.exe	891	1	1	1	0	2026-05-04	2026-05-04	0	1	197	197	No	No	No	No	Si	No	0	597d2bb32793
09 - cmd.exe	09	cmd.exe	1070	1	1	1	0	2026-04-04	2026-04-04	1	0	177	177	No	Si	No	No	No	No	0	3f650261e9ab
09 - cmd.exe	09	cmd.exe	1257	1	1	1	0	2026-01-21	2026-01-21	1	0	121	121	No	No	No	No	Si	No	0	5698bc5e87e3
09 - cmd.exe	09	cmd.exe	1264	1	1	1	0	2026-01-17	2026-01-17	1	0	118	118	No	No	No	No	Si	No	0	5b96433eb366
09 - cmd.exe	09	cmd.exe	1273	1	1	1	0	2026-01-11	2026-01-11	0	1	143	143	No	Si	No	Si	Si	No	0	9d39c2baae05
09 - cmd.exe	09	cmd.exe	1275	1	1	1	0	2026-01-10	2026-01-10	1	0	159	159	No	Si	No	Si	Si	No	0	95fcbc3986d8
09 - cmd.exe	09	cmd.exe	1307	1	1	1	0	2025-12-22	2025-12-22	1	0	150	150	No	Si	No	No	Si	No	0	04724f5ea99b
09 - cmd.exe	09	cmd.exe	1309	1	1	1	0	2025-12-14	2025-12-14	0	1	196	196	No	Si	No	No	No	No	0	e67ded5acb76
09 - cmd.exe	09	cmd.exe	1312	1	1	1	0	2025-12-11	2025-12-11	1	0	205	205	No	Si	No	No	No	No	0	546b73ddc540
09 - cmd.exe	09	cmd.exe	1313	1	1	1	0	2025-12-10	2025-12-10	1	0	205	205	No	Si	No	No	No	No	0	8e83c00d40ce
09 - cmd.exe	09	cmd.exe	1319	1	1	1	0	2025-12-06	2025-12-06	0	1	531	531	No	No	No	No	No	No	0	4daefb55c584
09 - cmd.exe	09	cmd.exe	1320	1	1	1	0	2025-12-06	2025-12-06	0	1	196	196	No	Si	No	No	No	No	0	6f7438f9fb45
09 - cmd.exe	09	cmd.exe	1325	1	1	1	0	2025-12-04	2025-12-04	0	1	228	228	No	Si	No	No	Si	No	0	8fa4bbfa014d
09 - cmd.exe	09	cmd.exe	1326	1	1	1	0	2025-12-04	2025-12-04	0	1	150	150	No	Si	No	No	No	No	0	c223c9252558
09 - cmd.exe	09	cmd.exe	1329	1	1	1	0	2025-11-29	2025-11-29	1	0	193	193	No	Si	No	No	No	No	0	49717606c8e6
09 - cmd.exe	09	cmd.exe	1333	1	1	1	0	2025-11-26	2025-11-26	1	0	196	196	No	Si	No	No	No	No	0	bc3b98747a1f
09 - cmd.exe	09	cmd.exe	1340	1	1	1	0	2025-11-21	2025-11-21	0	1	242	242	No	Si	No	No	Si	No	0	27ccf2d9d4fb
09 - cmd.exe	09	cmd.exe	1346	1	1	1	0	2025-11-14	2025-11-14	0	1	52	52	No	Si	No	No	No	No	0	32a7132e6e2b
09 - cmd.exe	09	cmd.exe	1350	1	1	1	0	2025-11-12	2025-11-12	0	1	202	202	No	Si	No	No	Si	No	0	1ecc904996dd
09 - cmd.exe	09	cmd.exe	1357	1	1	1	0	2025-11-03	2025-11-03	1	0	261	261	No	Si	No	No	No	No	0	c6010a0d9ae7
09 - cmd.exe	09	cmd.exe	1359	1	1	1	0	2025-11-01	2025-11-01	0	1	210	210	No	Si	No	No	Si	No	0	48b89857f123
09 - cmd.exe	09	cmd.exe	1361	1	1	1	0	2025-10-29	2025-10-29	0	1	126	126	No	Si	No	No	Si	No	0	1007289c8bac
09 - cmd.exe	09	cmd.exe	1362	1	1	1	0	2025-10-28	2025-10-28	1	0	261	261	No	Si	No	No	No	No	0	f0b0e8981606
09 - cmd.exe	09	cmd.exe	1367	1	1	1	0	2025-10-22	2025-10-22	0	1	144	144	No	Si	No	No	No	No	0	3e9cd9eff360
09 - cmd.exe	09	cmd.exe	1368	1	1	1	0	2025-10-22	2025-10-22	0	1	240	240	No	Si	No	No	Si	No	0	cd006bf549f4
09 - cmd.exe	09	cmd.exe	1370	1	1	1	0	2025-10-22	2025-10-22	0	1	313	313	No	No	Si	No	No	No	0	0e188ec4671b
09 - cmd.exe	09	cmd.exe	1374	1	1	1	0	2025-10-12	2025-10-12	0	1	109	109	No	No	No	No	No	No	0	7f84d97dcb85
09 - cmd.exe	09	cmd.exe	1408	1	1	1	0	2025-09-08	2025-09-08	0	1	246	246	No	Si	No	No	Si	No	0	e5aef999ba25
09 - cmd.exe	09	cmd.exe	1427	1	1	1	0	2025-08-30	2025-08-30	0	1	202	202	No	Si	No	No	Si	No	0	e55e6b392cc3
09 - cmd.exe	09	cmd.exe	1429	1	1	1	0	2025-08-29	2025-08-29	0	1	246	246	No	Si	No	No	Si	No	0	9df612efec8e
09 - cmd.exe	09	cmd.exe	1432	1	1	1	0	2025-08-28	2025-08-28	0	1	110	110	No	No	No	No	No	No	0	d210807e0e33
09 - cmd.exe	09	cmd.exe	1436	1	1	1	0	2025-08-28	2025-08-28	1	0	135	135	No	Si	No	Si	Si	No	0	45f9ca60dede
09 - cmd.exe	09	cmd.exe	1452	1	1	1	0	2025-08-26	2025-08-26	0	1	240	240	No	Si	No	No	Si	No	0	3f197b810ec5
09 - cmd.exe	09	cmd.exe	1459	1	1	1	0	2025-08-25	2025-08-25	0	1	200	200	No	Si	No	No	Si	No	0	8277e97bbe25
09 - cmd.exe	09	cmd.exe	1471	1	1	1	0	2025-08-22	2025-08-22	0	1	205	205	No	Si	No	No	No	No	0	72f7947ef0f0
09 - cmd.exe	09	cmd.exe	1483	1	1	1	0	2025-08-20	2025-08-20	1	0	225	225	No	No	No	Si	Si	No	0	9425200669d3
10 - script PowerShell embebido	10	script PowerShell embebido	6	121	121	121	0	2026-05-04	2026-05-05	9	112	144	144	No	No	Si	Si	No	No	0	37d4fc6d0069
10 - script PowerShell embebido	10	script PowerShell embebido	10	74	74	74	0	2026-07-22	2026-07-23	13	61	674	674	No	No	Si	No	Si	No	0	0da248db64cd
10 - script PowerShell embebido	10	script PowerShell embebido	15	35	35	35	0	2026-07-24	2026-07-25	4	31	674	674	No	No	Si	No	Si	No	0	b31b362fc777
10 - script PowerShell embebido	10	script PowerShell embebido	22	24	24	24	0	2026-07-24	2026-07-24	5	19	670	670	No	No	Si	No	Si	No	0	4f8407088b3f
10 - script PowerShell embebido	10	script PowerShell embebido	37	12	12	12	0	2026-03-28	2026-07-26	12	0	186	186	No	No	No	Si	No	No	0	985eba8f1e8e
10 - script PowerShell embebido	10	script PowerShell embebido	79	4	4	4	0	2025-11-04	2025-11-18	0	4	203	203	No	Si	Si	Si	No	No	0	c0d126bde790
10 - script PowerShell embebido	10	script PowerShell embebido	82	3	3	3	0	2026-07-21	2026-07-21	2	1	670	670	No	No	Si	No	Si	No	0	81b88c95c431
10 - script PowerShell embebido	10	script PowerShell embebido	83	3	3	3	0	2025-12-04	2026-05-21	2	1	43	43	No	No	No	No	No	No	0	05267d2818fe
10 - script PowerShell embebido	10	script PowerShell embebido	85	3	3	2	1	2026-03-24	2026-03-25	3	0	22	22	No	No	No	No	No	No	0	f834a46bd638
10 - script PowerShell embebido	10	script PowerShell embebido	93	3	3	3	0	2025-12-08	2025-12-12	0	3	179	179	No	Si	Si	Si	No	No	0	b1ad8f84db84
10 - script PowerShell embebido	10	script PowerShell embebido	117	2	2	2	0	2026-06-17	2026-06-17	2	0	206	206	No	No	No	Si	No	No	0	84129278ab49
10 - script PowerShell embebido	10	script PowerShell embebido	118	2	2	2	0	2026-05-30	2026-05-30	0	2	2151	2151	No	No	No	Si	No	No	0	8db67dc6d121
10 - script PowerShell embebido	10	script PowerShell embebido	128	2	2	2	0	2026-03-21	2026-03-21	0	2	4110	4110	No	No	No	Si	No	No	0	f96a24d07929
10 - script PowerShell embebido	10	script PowerShell embebido	158	2	2	2	0	2025-08-20	2025-08-25	0	2	49	49	No	Si	No	Si	No	No	0	578cf2972033
10 - script PowerShell embebido	10	script PowerShell embebido	169	1	1	1	0	2026-07-28	2026-07-28	1	0	116	116	No	Si	No	Si	No	No	0	81ebac3d16b2
10 - script PowerShell embebido	10	script PowerShell embebido	237	1	1	1	0	2026-07-23	2026-07-23	0	1	666	666	No	No	Si	No	Si	No	0	3f2c51d6cb84
10 - script PowerShell embebido	10	script PowerShell embebido	255	1	1	1	0	2026-07-17	2026-07-17	0	1	674	674	No	No	Si	No	Si	No	0	bfbea440beb4
10 - script PowerShell embebido	10	script PowerShell embebido	319	1	1	1	0	2026-07-01	2026-07-01	1	0	844	844	No	Si	No	No	Si	No	0	d29729ecf1eb
10 - script PowerShell embebido	10	script PowerShell embebido	360	1	1	1	0	2026-06-26	2026-06-26	1	0	838	838	No	Si	No	No	Si	No	0	3a8775d2c618
10 - script PowerShell embebido	10	script PowerShell embebido	389	1	1	1	0	2026-06-24	2026-06-24	0	1	842	842	No	Si	No	No	Si	No	0	64b404f238c8
10 - script PowerShell embebido	10	script PowerShell embebido	390	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	a641b82f4ddb
10 - script PowerShell embebido	10	script PowerShell embebido	391	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	8a84621c9639
10 - script PowerShell embebido	10	script PowerShell embebido	393	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	1fbcc83d3484
10 - script PowerShell embebido	10	script PowerShell embebido	394	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	b2a1b5730000
10 - script PowerShell embebido	10	script PowerShell embebido	395	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	35460339db5e
10 - script PowerShell embebido	10	script PowerShell embebido	396	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	53050d803093
10 - script PowerShell embebido	10	script PowerShell embebido	397	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	c2acd8c6ae7f
10 - script PowerShell embebido	10	script PowerShell embebido	398	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	00a8290fe8ba
10 - script PowerShell embebido	10	script PowerShell embebido	400	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	26bfa8741e17
10 - script PowerShell embebido	10	script PowerShell embebido	401	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	71d1d052dcaa
10 - script PowerShell embebido	10	script PowerShell embebido	403	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	0b0d393ebe78
10 - script PowerShell embebido	10	script PowerShell embebido	404	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	22bea0f2ea6a
10 - script PowerShell embebido	10	script PowerShell embebido	405	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	8b3ac0bf2e16
10 - script PowerShell embebido	10	script PowerShell embebido	406	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	3a52496bad5d
10 - script PowerShell embebido	10	script PowerShell embebido	407	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	1ddd3f5b81ee
10 - script PowerShell embebido	10	script PowerShell embebido	408	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	00cf1e70d9e8
10 - script PowerShell embebido	10	script PowerShell embebido	409	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	1f1073748f38
10 - script PowerShell embebido	10	script PowerShell embebido	410	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	67594fcf3f3f
10 - script PowerShell embebido	10	script PowerShell embebido	411	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	f53768bdf01c
10 - script PowerShell embebido	10	script PowerShell embebido	412	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	59c17238d37d
10 - script PowerShell embebido	10	script PowerShell embebido	413	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	a7a8a17588fb
10 - script PowerShell embebido	10	script PowerShell embebido	414	1	1	1	0	2026-06-24	2026-06-24	0	1	842	842	No	Si	No	No	Si	No	0	d76ea7499c87
10 - script PowerShell embebido	10	script PowerShell embebido	415	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	4e1c4d2f13fb
10 - script PowerShell embebido	10	script PowerShell embebido	416	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	f8a2f4d901de
10 - script PowerShell embebido	10	script PowerShell embebido	417	1	1	1	0	2026-06-24	2026-06-24	1	0	839	839	No	Si	No	No	Si	No	0	09545a1c67ce
10 - script PowerShell embebido	10	script PowerShell embebido	418	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	d98bb00f780f
10 - script PowerShell embebido	10	script PowerShell embebido	419	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	ae3850fdba6f
10 - script PowerShell embebido	10	script PowerShell embebido	420	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	b1fe2892ae3f
10 - script PowerShell embebido	10	script PowerShell embebido	421	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	192361cb941a
10 - script PowerShell embebido	10	script PowerShell embebido	422	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	1b25508f9758
10 - script PowerShell embebido	10	script PowerShell embebido	423	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	3157d429d0d1
10 - script PowerShell embebido	10	script PowerShell embebido	424	1	1	1	0	2026-06-24	2026-06-24	1	0	842	842	No	Si	No	No	Si	No	0	71198e8fbe57
10 - script PowerShell embebido	10	script PowerShell embebido	428	1	1	1	0	2026-06-24	2026-06-24	1	0	835	835	No	Si	No	No	Si	No	0	bef52991c516
10 - script PowerShell embebido	10	script PowerShell embebido	429	1	1	1	0	2026-06-24	2026-06-24	1	0	835	835	No	Si	No	No	Si	No	0	36099b8e14ea
10 - script PowerShell embebido	10	script PowerShell embebido	432	1	1	1	0	2026-06-23	2026-06-23	1	0	827	827	No	Si	No	No	Si	No	0	45301bf77e6f
10 - script PowerShell embebido	10	script PowerShell embebido	434	1	1	1	0	2026-06-23	2026-06-23	1	0	835	835	No	Si	No	No	Si	No	0	2918968109c7
10 - script PowerShell embebido	10	script PowerShell embebido	435	1	1	1	0	2026-06-23	2026-06-23	1	0	835	835	No	Si	No	No	Si	No	0	4efd5b6ec8cb
10 - script PowerShell embebido	10	script PowerShell embebido	436	1	1	1	0	2026-06-23	2026-06-23	1	0	835	835	No	Si	No	No	Si	No	0	07978668d329
10 - script PowerShell embebido	10	script PowerShell embebido	440	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	f56bb92e6c41
10 - script PowerShell embebido	10	script PowerShell embebido	441	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	12b0e1fd935c
10 - script PowerShell embebido	10	script PowerShell embebido	442	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	176b3b907923
10 - script PowerShell embebido	10	script PowerShell embebido	443	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	e438069415f4
10 - script PowerShell embebido	10	script PowerShell embebido	444	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	6a2f40083524
10 - script PowerShell embebido	10	script PowerShell embebido	445	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	9b1417716706
10 - script PowerShell embebido	10	script PowerShell embebido	446	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	ff68a7b1f743
10 - script PowerShell embebido	10	script PowerShell embebido	447	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	eecdf6e53c28
10 - script PowerShell embebido	10	script PowerShell embebido	448	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	3d4d73c547d4
10 - script PowerShell embebido	10	script PowerShell embebido	449	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	436dcfa7a06b
10 - script PowerShell embebido	10	script PowerShell embebido	450	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	79c710b98c33
10 - script PowerShell embebido	10	script PowerShell embebido	451	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	e92b29114d12
10 - script PowerShell embebido	10	script PowerShell embebido	452	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	e28ae76f46f2
10 - script PowerShell embebido	10	script PowerShell embebido	453	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	1d8a720e7ff8
10 - script PowerShell embebido	10	script PowerShell embebido	454	1	1	1	0	2026-06-22	2026-06-22	0	1	833	833	No	Si	No	No	Si	No	0	ad81d87fceb8
10 - script PowerShell embebido	10	script PowerShell embebido	455	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	c2da0e6d5001
10 - script PowerShell embebido	10	script PowerShell embebido	456	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	bcbd65f9900d
10 - script PowerShell embebido	10	script PowerShell embebido	457	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	85b00f2a02ed
10 - script PowerShell embebido	10	script PowerShell embebido	458	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	cd6bb47776ec
10 - script PowerShell embebido	10	script PowerShell embebido	459	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	189cf8cec95b
10 - script PowerShell embebido	10	script PowerShell embebido	460	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	893edd8d3015
10 - script PowerShell embebido	10	script PowerShell embebido	461	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	f791bf1899f8
10 - script PowerShell embebido	10	script PowerShell embebido	462	1	1	1	0	2026-06-22	2026-06-22	1	0	833	833	No	Si	No	No	Si	No	0	4542809a0f2e
10 - script PowerShell embebido	10	script PowerShell embebido	463	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	86cece90c1d5
10 - script PowerShell embebido	10	script PowerShell embebido	464	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	4a40c020b40d
10 - script PowerShell embebido	10	script PowerShell embebido	465	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	5d254c6515b8
10 - script PowerShell embebido	10	script PowerShell embebido	466	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	d5986483caf7
10 - script PowerShell embebido	10	script PowerShell embebido	467	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	983e17c452d6
10 - script PowerShell embebido	10	script PowerShell embebido	468	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	2c21e802a1f8
10 - script PowerShell embebido	10	script PowerShell embebido	469	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	ff98528e5789
10 - script PowerShell embebido	10	script PowerShell embebido	470	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	6a9533164a09
10 - script PowerShell embebido	10	script PowerShell embebido	471	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	1a19a8308cee
10 - script PowerShell embebido	10	script PowerShell embebido	472	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	a8648f03e645
10 - script PowerShell embebido	10	script PowerShell embebido	473	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	f1ab6f38a0d0
10 - script PowerShell embebido	10	script PowerShell embebido	474	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	0e26307929da
10 - script PowerShell embebido	10	script PowerShell embebido	475	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	1757b4702e8c
10 - script PowerShell embebido	10	script PowerShell embebido	476	1	1	1	0	2026-06-22	2026-06-22	0	1	833	833	No	Si	No	No	Si	No	0	e879980c8572
10 - script PowerShell embebido	10	script PowerShell embebido	477	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	24c418055555
10 - script PowerShell embebido	10	script PowerShell embebido	478	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	35fe762fa011
10 - script PowerShell embebido	10	script PowerShell embebido	479	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	c07983adb715
10 - script PowerShell embebido	10	script PowerShell embebido	480	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	4c9cf17d2665
10 - script PowerShell embebido	10	script PowerShell embebido	481	1	1	1	0	2026-06-22	2026-06-22	1	0	835	835	No	Si	No	No	Si	No	0	2768aaf94fad
10 - script PowerShell embebido	10	script PowerShell embebido	482	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	d99b095d0a1c
10 - script PowerShell embebido	10	script PowerShell embebido	483	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	167537c7268b
10 - script PowerShell embebido	10	script PowerShell embebido	484	1	1	1	0	2026-06-22	2026-06-22	1	0	833	833	No	Si	No	No	Si	No	0	1147d21ef831
10 - script PowerShell embebido	10	script PowerShell embebido	485	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	870efa9540f4
10 - script PowerShell embebido	10	script PowerShell embebido	486	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	165251081b13
10 - script PowerShell embebido	10	script PowerShell embebido	487	1	1	1	0	2026-06-22	2026-06-22	0	1	835	835	No	Si	No	No	Si	No	0	ea6215ee814c
10 - script PowerShell embebido	10	script PowerShell embebido	488	1	1	1	0	2026-06-22	2026-06-22	1	0	833	833	No	Si	No	No	Si	No	0	8898b7754635
10 - script PowerShell embebido	10	script PowerShell embebido	489	1	1	1	0	2026-06-22	2026-06-22	0	1	881	881	No	Si	No	No	Si	No	0	847bddb04364
10 - script PowerShell embebido	10	script PowerShell embebido	490	1	1	1	0	2026-06-22	2026-06-22	1	0	881	881	No	Si	No	No	Si	No	0	778c307367c3
10 - script PowerShell embebido	10	script PowerShell embebido	491	1	1	1	0	2026-06-22	2026-06-22	1	0	881	881	No	Si	No	No	Si	No	0	0cff39e0d2b1
10 - script PowerShell embebido	10	script PowerShell embebido	492	1	1	1	0	2026-06-22	2026-06-22	1	0	881	881	No	Si	No	No	Si	No	0	cde3671346e9
10 - script PowerShell embebido	10	script PowerShell embebido	493	1	1	1	0	2026-06-22	2026-06-22	1	0	881	881	No	Si	No	No	Si	No	0	2a6e7a596473
10 - script PowerShell embebido	10	script PowerShell embebido	495	1	1	1	0	2026-06-22	2026-06-22	1	0	881	881	No	Si	No	No	Si	No	0	a4a589639ee1
10 - script PowerShell embebido	10	script PowerShell embebido	496	1	1	1	0	2026-06-22	2026-06-22	0	1	881	881	No	Si	No	No	Si	No	0	0d4fcb0b5834
10 - script PowerShell embebido	10	script PowerShell embebido	497	1	1	1	0	2026-06-22	2026-06-22	1	0	881	881	No	Si	No	No	Si	No	0	7f71d9dc2c4f
10 - script PowerShell embebido	10	script PowerShell embebido	498	1	1	1	0	2026-06-22	2026-06-22	1	0	881	881	No	Si	No	No	Si	No	0	db0072eff13b
10 - script PowerShell embebido	10	script PowerShell embebido	499	1	1	1	0	2026-06-22	2026-06-22	1	0	881	881	No	Si	No	No	Si	No	0	c7a624a4334b
10 - script PowerShell embebido	10	script PowerShell embebido	500	1	1	1	0	2026-06-22	2026-06-22	0	1	881	881	No	Si	No	No	Si	No	0	4fae013bf277
10 - script PowerShell embebido	10	script PowerShell embebido	501	1	1	1	0	2026-06-22	2026-06-22	1	0	881	881	No	Si	No	No	Si	No	0	7267923629c5
10 - script PowerShell embebido	10	script PowerShell embebido	502	1	1	1	0	2026-06-22	2026-06-22	1	0	881	881	No	Si	No	No	Si	No	0	f8c1d32e2cb1
10 - script PowerShell embebido	10	script PowerShell embebido	503	1	1	1	0	2026-06-22	2026-06-22	0	1	881	881	No	Si	No	No	Si	No	0	dd489c053bbe
10 - script PowerShell embebido	10	script PowerShell embebido	504	1	1	1	0	2026-06-22	2026-06-22	1	0	881	881	No	Si	No	No	Si	No	0	bbb043ee213d
10 - script PowerShell embebido	10	script PowerShell embebido	505	1	1	1	0	2026-06-22	2026-06-22	0	1	881	881	No	Si	No	No	Si	No	0	2dccce4ddf02
10 - script PowerShell embebido	10	script PowerShell embebido	506	1	1	1	0	2026-06-22	2026-06-22	1	0	881	881	No	Si	No	No	Si	No	0	b1334816c789
10 - script PowerShell embebido	10	script PowerShell embebido	507	1	1	1	0	2026-06-22	2026-06-22	0	1	881	881	No	Si	No	No	Si	No	0	58b99f336ae4
10 - script PowerShell embebido	10	script PowerShell embebido	508	1	1	1	0	2026-06-22	2026-06-22	0	1	881	881	No	Si	No	No	Si	No	0	1bfcc93e6d95
10 - script PowerShell embebido	10	script PowerShell embebido	509	1	1	1	0	2026-06-22	2026-06-22	0	1	881	881	No	Si	No	No	Si	No	0	07b75e3608cc
10 - script PowerShell embebido	10	script PowerShell embebido	510	1	1	1	0	2026-06-22	2026-06-22	1	0	879	879	No	Si	No	No	Si	No	0	f4e696143cae
10 - script PowerShell embebido	10	script PowerShell embebido	511	1	1	1	0	2026-06-22	2026-06-22	0	1	881	881	No	Si	No	No	Si	No	0	c5ca3099c9a1
10 - script PowerShell embebido	10	script PowerShell embebido	512	1	1	1	0	2026-06-22	2026-06-22	1	0	881	881	No	Si	No	No	Si	No	0	a3b6daac5b01
10 - script PowerShell embebido	10	script PowerShell embebido	513	1	1	1	0	2026-06-22	2026-06-22	0	1	881	881	No	Si	No	No	Si	No	0	5e6aebd2cbc3
10 - script PowerShell embebido	10	script PowerShell embebido	515	1	1	1	0	2026-06-22	2026-06-22	0	1	881	881	No	Si	No	No	Si	No	0	5bde237ff1bd
10 - script PowerShell embebido	10	script PowerShell embebido	518	1	1	1	0	2026-06-19	2026-06-19	1	0	880	880	No	Si	No	No	Si	No	0	bbcfdf253a19
10 - script PowerShell embebido	10	script PowerShell embebido	520	1	1	1	0	2026-06-19	2026-06-19	0	1	903	903	No	Si	No	Si	No	No	0	c2adb9cc4ce7
10 - script PowerShell embebido	10	script PowerShell embebido	523	1	1	1	0	2026-06-18	2026-06-18	0	1	885	885	No	No	No	Si	No	No	0	9d77338e2f7c
10 - script PowerShell embebido	10	script PowerShell embebido	524	1	1	1	0	2026-06-17	2026-06-17	0	1	883	883	No	Si	No	Si	No	No	0	9361cdcb97dc
10 - script PowerShell embebido	10	script PowerShell embebido	525	1	1	1	0	2026-06-17	2026-06-17	1	0	691	691	No	No	No	Si	No	No	0	7cdab5c1b876
10 - script PowerShell embebido	10	script PowerShell embebido	526	1	1	1	0	2026-06-16	2026-06-16	1	0	878	878	No	Si	No	No	Si	No	0	f6ea11fc5dd3
10 - script PowerShell embebido	10	script PowerShell embebido	527	1	1	1	0	2026-06-16	2026-06-16	0	1	909	909	No	Si	No	Si	No	No	0	cb057ca8ba10
10 - script PowerShell embebido	10	script PowerShell embebido	528	1	1	1	0	2026-06-16	2026-06-16	1	0	713	713	No	No	No	Si	No	No	0	bd4e2851e4ba
10 - script PowerShell embebido	10	script PowerShell embebido	529	1	1	1	0	2026-06-16	2026-06-16	0	1	2094	2094	No	No	No	Si	No	No	0	359e94e04780
10 - script PowerShell embebido	10	script PowerShell embebido	530	1	1	1	0	2026-06-15	2026-06-15	0	1	813	813	No	Si	No	Si	No	No	0	efbc233dc9e6
10 - script PowerShell embebido	10	script PowerShell embebido	533	1	1	1	0	2026-06-14	2026-06-14	0	1	628	628	No	No	No	Si	No	No	0	ac8a2f1a0e23
10 - script PowerShell embebido	10	script PowerShell embebido	534	1	1	1	0	2026-06-14	2026-06-14	0	1	865	865	No	No	No	Si	No	No	0	78c29d20c75f
10 - script PowerShell embebido	10	script PowerShell embebido	535	1	1	1	0	2026-06-14	2026-06-14	1	0	919	919	No	Si	No	Si	No	No	0	9595103d358f
10 - script PowerShell embebido	10	script PowerShell embebido	536	1	1	1	0	2026-06-13	2026-06-13	0	1	752	752	No	No	No	Si	No	No	0	3927d3d10347
10 - script PowerShell embebido	10	script PowerShell embebido	537	1	1	1	0	2026-06-13	2026-06-13	1	0	875	875	No	Si	No	No	Si	No	0	bbb4f2bc2502
10 - script PowerShell embebido	10	script PowerShell embebido	538	1	1	1	0	2026-06-12	2026-06-12	0	1	2441	2441	No	No	No	Si	No	No	0	3a3eac06b542
10 - script PowerShell embebido	10	script PowerShell embebido	539	1	1	1	0	2026-06-12	2026-06-12	0	1	2506	2506	No	No	No	Si	No	No	0	2567a1efeeb4
10 - script PowerShell embebido	10	script PowerShell embebido	541	1	1	1	0	2026-06-11	2026-06-11	0	1	2475	2475	No	No	No	Si	No	No	0	1187e63e74a6
10 - script PowerShell embebido	10	script PowerShell embebido	542	1	1	1	0	2026-06-11	2026-06-11	0	1	2536	2536	No	No	No	Si	No	No	0	7bb116246a8c
10 - script PowerShell embebido	10	script PowerShell embebido	543	1	1	1	0	2026-06-10	2026-06-10	0	1	2450	2450	No	No	No	Si	No	No	0	1ae417a5a86b
10 - script PowerShell embebido	10	script PowerShell embebido	544	1	1	1	0	2026-06-10	2026-06-10	1	0	26	26	No	No	No	No	No	No	0	fb60143ed972
10 - script PowerShell embebido	10	script PowerShell embebido	545	1	1	1	0	2026-06-10	2026-06-10	0	1	2538	2538	No	No	No	Si	No	No	0	bc72283c0bf5
10 - script PowerShell embebido	10	script PowerShell embebido	546	1	1	1	0	2026-06-10	2026-06-10	0	1	2505	2505	No	No	No	Si	No	No	0	cd49ec2f101c
10 - script PowerShell embebido	10	script PowerShell embebido	547	1	1	1	0	2026-06-10	2026-06-10	0	1	2490	2490	No	No	No	Si	No	No	0	a5ac8daf99ae
10 - script PowerShell embebido	10	script PowerShell embebido	549	1	1	1	0	2026-06-09	2026-06-09	0	1	2564	2564	No	No	No	Si	No	No	0	73f02cddf456
10 - script PowerShell embebido	10	script PowerShell embebido	550	1	1	1	0	2026-06-09	2026-06-09	1	0	2499	2499	No	No	No	Si	No	No	0	d8b6c1b61e57
10 - script PowerShell embebido	10	script PowerShell embebido	551	1	1	1	0	2026-06-09	2026-06-09	1	0	2478	2478	No	No	No	Si	No	No	0	c95fc1464a39
10 - script PowerShell embebido	10	script PowerShell embebido	553	1	1	1	0	2026-06-08	2026-06-08	0	1	2539	2539	No	No	No	Si	No	No	0	28a37fa4bd12
10 - script PowerShell embebido	10	script PowerShell embebido	554	1	1	1	0	2026-06-08	2026-06-08	1	0	2498	2498	No	No	No	Si	No	No	0	562a9c2aea4d
10 - script PowerShell embebido	10	script PowerShell embebido	555	1	1	1	0	2026-06-08	2026-06-08	1	0	2478	2478	No	No	No	Si	No	No	0	c7c2a8d6d3ce
10 - script PowerShell embebido	10	script PowerShell embebido	557	1	1	1	0	2026-06-08	2026-06-08	0	1	2529	2529	No	No	No	Si	No	No	0	70a36c97ea1a
10 - script PowerShell embebido	10	script PowerShell embebido	558	1	1	1	0	2026-06-08	2026-06-08	1	0	2484	2484	No	No	No	Si	No	No	0	902e4203f2a7
10 - script PowerShell embebido	10	script PowerShell embebido	559	1	1	1	0	2026-06-08	2026-06-08	1	0	2500	2500	No	No	No	Si	No	No	0	154b48120ec2
10 - script PowerShell embebido	10	script PowerShell embebido	560	1	1	1	0	2026-06-08	2026-06-08	1	0	2574	2574	No	No	No	Si	No	No	0	d97164c8cdbd
10 - script PowerShell embebido	10	script PowerShell embebido	561	1	1	1	0	2026-06-08	2026-06-08	1	0	2435	2435	No	No	No	Si	No	No	0	7eb904f6a64c
10 - script PowerShell embebido	10	script PowerShell embebido	562	1	1	1	0	2026-06-08	2026-06-08	1	0	2528	2528	No	No	No	Si	No	No	0	a0a20350617b
10 - script PowerShell embebido	10	script PowerShell embebido	563	1	1	1	0	2026-06-08	2026-06-08	0	1	2486	2486	No	No	No	Si	No	No	0	a91585849950
10 - script PowerShell embebido	10	script PowerShell embebido	564	1	1	1	0	2026-06-08	2026-06-08	0	1	2502	2502	No	No	No	Si	No	No	0	aafcf718dc80
10 - script PowerShell embebido	10	script PowerShell embebido	565	1	1	1	0	2026-06-08	2026-06-08	0	1	2477	2477	No	No	No	Si	No	No	0	5d0791700900
10 - script PowerShell embebido	10	script PowerShell embebido	566	1	1	1	0	2026-06-08	2026-06-08	1	0	2541	2541	No	No	No	Si	No	No	0	b84725eb020c
10 - script PowerShell embebido	10	script PowerShell embebido	567	1	1	1	0	2026-06-08	2026-06-08	0	1	2469	2469	No	No	No	Si	No	No	0	dc4944f5a52d
10 - script PowerShell embebido	10	script PowerShell embebido	568	1	1	1	0	2026-06-08	2026-06-08	0	1	2495	2495	No	No	No	Si	No	No	0	a37b0d219a8d
10 - script PowerShell embebido	10	script PowerShell embebido	569	1	1	1	0	2026-06-08	2026-06-08	1	0	2451	2451	No	No	No	Si	No	No	0	237b0675caab
10 - script PowerShell embebido	10	script PowerShell embebido	570	1	1	1	0	2026-06-08	2026-06-08	1	0	2524	2524	No	No	No	Si	No	No	0	e9eb2b009a7a
10 - script PowerShell embebido	10	script PowerShell embebido	571	1	1	1	0	2026-06-08	2026-06-08	0	1	2485	2485	No	No	No	Si	No	No	0	ccd0f7c76b24
10 - script PowerShell embebido	10	script PowerShell embebido	572	1	1	1	0	2026-06-08	2026-06-08	1	0	2547	2547	No	No	No	Si	No	No	0	12463c93a696
10 - script PowerShell embebido	10	script PowerShell embebido	573	1	1	1	0	2026-06-08	2026-06-08	0	1	2454	2454	No	No	No	Si	No	No	0	8188a4868e00
10 - script PowerShell embebido	10	script PowerShell embebido	574	1	1	1	0	2026-06-08	2026-06-08	1	0	2516	2516	No	No	No	Si	No	No	0	ec0c50c5d1e6
10 - script PowerShell embebido	10	script PowerShell embebido	575	1	1	1	0	2026-06-08	2026-06-08	1	0	2471	2471	No	No	No	Si	No	No	0	a7377123dc79
10 - script PowerShell embebido	10	script PowerShell embebido	576	1	1	1	0	2026-06-08	2026-06-08	0	1	2532	2532	No	No	No	Si	No	No	0	6b93a4bf1d23
10 - script PowerShell embebido	10	script PowerShell embebido	577	1	1	1	0	2026-06-08	2026-06-08	0	1	2491	2491	No	No	No	Si	No	No	0	fd62790a4a12
10 - script PowerShell embebido	10	script PowerShell embebido	578	1	1	1	0	2026-06-08	2026-06-08	0	1	2556	2556	No	No	No	Si	No	No	0	fe1db77b33a6
10 - script PowerShell embebido	10	script PowerShell embebido	579	1	1	1	0	2026-06-08	2026-06-08	0	1	2549	2549	No	No	No	Si	No	No	0	eadaa4f6965c
10 - script PowerShell embebido	10	script PowerShell embebido	580	1	1	1	0	2026-06-08	2026-06-08	0	1	2521	2521	No	No	No	Si	No	No	0	cbd9d307ad3b
10 - script PowerShell embebido	10	script PowerShell embebido	581	1	1	1	0	2026-06-08	2026-06-08	1	0	2531	2531	No	No	No	Si	No	No	0	a6984a81cb6d
10 - script PowerShell embebido	10	script PowerShell embebido	582	1	1	1	0	2026-06-08	2026-06-08	1	0	2514	2514	No	No	No	Si	No	No	0	9d665fa29495
10 - script PowerShell embebido	10	script PowerShell embebido	583	1	1	1	0	2026-06-08	2026-06-08	1	0	2583	2583	No	No	No	Si	No	No	0	de3781e02822
10 - script PowerShell embebido	10	script PowerShell embebido	584	1	1	1	0	2026-06-08	2026-06-08	0	1	2523	2523	No	No	No	Si	No	No	0	8e0257fbbcec
10 - script PowerShell embebido	10	script PowerShell embebido	585	1	1	1	0	2026-06-08	2026-06-08	0	1	2539	2539	No	No	No	Si	No	No	0	a9f39a055f51
10 - script PowerShell embebido	10	script PowerShell embebido	586	1	1	1	0	2026-06-08	2026-06-08	1	0	2578	2578	No	No	No	Si	No	No	0	6ff7c36fcf72
10 - script PowerShell embebido	10	script PowerShell embebido	587	1	1	1	0	2026-06-08	2026-06-08	0	1	2582	2582	No	No	No	Si	No	No	0	d41324858611
10 - script PowerShell embebido	10	script PowerShell embebido	588	1	1	1	0	2026-06-08	2026-06-08	1	0	2463	2463	No	No	No	Si	No	No	0	e2ce0896eb9b
10 - script PowerShell embebido	10	script PowerShell embebido	589	1	1	1	0	2026-06-08	2026-06-08	1	0	2470	2470	No	No	No	Si	No	No	0	474cb42852aa
10 - script PowerShell embebido	10	script PowerShell embebido	590	1	1	1	0	2026-06-08	2026-06-08	1	0	2504	2504	No	No	No	Si	No	No	0	2d1ec8bdf63b
10 - script PowerShell embebido	10	script PowerShell embebido	591	1	1	1	0	2026-06-08	2026-06-08	0	1	2518	2518	No	No	No	Si	No	No	0	4547be8c1800
10 - script PowerShell embebido	10	script PowerShell embebido	592	1	1	1	0	2026-06-08	2026-06-08	1	0	2564	2564	No	No	No	Si	No	No	0	9b01024cc95b
10 - script PowerShell embebido	10	script PowerShell embebido	593	1	1	1	0	2026-06-08	2026-06-08	1	0	2514	2514	No	No	No	Si	No	No	0	a83f805eeef0
10 - script PowerShell embebido	10	script PowerShell embebido	594	1	1	1	0	2026-06-08	2026-06-08	1	0	2551	2551	No	No	No	Si	No	No	0	95d071f2f739
10 - script PowerShell embebido	10	script PowerShell embebido	595	1	1	1	0	2026-06-08	2026-06-08	0	1	2493	2493	No	No	No	Si	No	No	0	384a70c385fe
10 - script PowerShell embebido	10	script PowerShell embebido	597	1	1	1	0	2026-06-08	2026-06-08	0	1	2515	2515	No	No	No	Si	No	No	0	a7222acadfc0
10 - script PowerShell embebido	10	script PowerShell embebido	598	1	1	1	0	2026-06-08	2026-06-08	0	1	2516	2516	No	No	No	Si	No	No	0	3306c995d79e
10 - script PowerShell embebido	10	script PowerShell embebido	599	1	1	1	0	2026-06-08	2026-06-08	1	0	2591	2591	No	No	No	Si	No	No	0	684d448293e2
10 - script PowerShell embebido	10	script PowerShell embebido	600	1	1	1	0	2026-06-08	2026-06-08	1	0	2475	2475	No	No	No	Si	No	No	0	397c9ab70d09
10 - script PowerShell embebido	10	script PowerShell embebido	603	1	1	1	0	2026-05-31	2026-05-31	0	1	2169	2169	No	No	No	Si	No	No	0	81565c7cb6af
10 - script PowerShell embebido	10	script PowerShell embebido	604	1	1	1	0	2026-05-31	2026-05-31	0	1	4112	4112	No	No	No	Si	No	No	0	b71f7341a3c4
10 - script PowerShell embebido	10	script PowerShell embebido	605	1	1	1	0	2026-05-30	2026-05-30	0	1	2164	2164	No	No	No	Si	No	No	0	65861ee27661
10 - script PowerShell embebido	10	script PowerShell embebido	607	1	1	1	0	2026-05-30	2026-05-30	0	1	2134	2134	No	No	No	Si	No	No	0	0a565ab70dbf
10 - script PowerShell embebido	10	script PowerShell embebido	608	1	1	1	0	2026-05-30	2026-05-30	0	1	2182	2182	No	No	No	Si	No	No	0	ee16e24bfb59
10 - script PowerShell embebido	10	script PowerShell embebido	610	1	1	1	0	2026-05-27	2026-05-27	0	1	1897	1897	No	No	Si	Si	No	No	0	e5ad711dcb5b
10 - script PowerShell embebido	10	script PowerShell embebido	612	1	1	1	0	2026-05-27	2026-05-27	1	0	2040	2040	No	No	Si	Si	No	No	0	f98787b4adcb
10 - script PowerShell embebido	10	script PowerShell embebido	613	1	1	1	0	2026-05-27	2026-05-27	0	1	1893	1893	No	No	Si	Si	No	No	0	32fdc296f52a
10 - script PowerShell embebido	10	script PowerShell embebido	616	1	1	1	0	2026-05-25	2026-05-25	0	1	2042	2042	No	No	Si	Si	No	No	0	12d8a875d704
10 - script PowerShell embebido	10	script PowerShell embebido	617	1	1	1	0	2026-05-24	2026-05-24	1	0	1893	1893	No	No	Si	Si	No	No	0	6202e55f570e
10 - script PowerShell embebido	10	script PowerShell embebido	618	1	1	1	0	2026-05-24	2026-05-24	0	1	2054	2054	No	No	Si	Si	No	No	0	0e6324256207
10 - script PowerShell embebido	10	script PowerShell embebido	619	1	1	1	0	2026-05-24	2026-05-24	0	1	1892	1892	No	No	Si	Si	No	No	0	ac83323676e2
10 - script PowerShell embebido	10	script PowerShell embebido	620	1	1	1	0	2026-05-24	2026-05-24	1	0	2038	2038	No	No	Si	Si	No	No	0	9ccbd0c63970
10 - script PowerShell embebido	10	script PowerShell embebido	621	1	1	1	0	2026-05-24	2026-05-24	0	1	2080	2080	No	No	Si	Si	No	No	0	9f9225da8df3
10 - script PowerShell embebido	10	script PowerShell embebido	622	1	1	1	0	2026-05-24	2026-05-24	1	0	1923	1923	No	No	Si	Si	No	No	0	51e955c90f29
10 - script PowerShell embebido	10	script PowerShell embebido	623	1	1	1	0	2026-05-24	2026-05-24	1	0	2049	2049	No	No	Si	Si	No	No	0	f596d121ea58
10 - script PowerShell embebido	10	script PowerShell embebido	624	1	1	1	0	2026-05-24	2026-05-24	1	0	2065	2065	No	No	Si	Si	No	No	0	091b79491a99
10 - script PowerShell embebido	10	script PowerShell embebido	625	1	1	1	0	2026-05-24	2026-05-24	0	1	2094	2094	No	No	Si	Si	No	No	0	1e06f5ff17ca
10 - script PowerShell embebido	10	script PowerShell embebido	626	1	1	1	0	2026-05-24	2026-05-24	0	1	1893	1893	No	No	Si	Si	No	No	0	38e74d1dec18
10 - script PowerShell embebido	10	script PowerShell embebido	627	1	1	1	0	2026-05-24	2026-05-24	0	1	1921	1921	No	No	Si	Si	No	No	0	f6aad9712f34
10 - script PowerShell embebido	10	script PowerShell embebido	628	1	1	1	0	2026-05-24	2026-05-24	1	0	1902	1902	No	No	Si	Si	No	No	0	1b8771c622a9
10 - script PowerShell embebido	10	script PowerShell embebido	629	1	1	1	0	2026-05-24	2026-05-24	0	1	2040	2040	No	No	Si	Si	No	No	0	f9bf24d53b6b
10 - script PowerShell embebido	10	script PowerShell embebido	630	1	1	1	0	2026-05-23	2026-05-23	0	1	2077	2077	No	No	Si	Si	No	No	0	62ded84e78ea
10 - script PowerShell embebido	10	script PowerShell embebido	631	1	1	1	0	2026-05-22	2026-05-22	0	1	2106	2106	No	No	Si	Si	No	No	0	8fdc2d3a2c78
10 - script PowerShell embebido	10	script PowerShell embebido	632	1	1	1	0	2026-05-22	2026-05-22	1	0	1907	1907	No	No	Si	Si	No	No	0	3eb60c117cac
10 - script PowerShell embebido	10	script PowerShell embebido	633	1	1	1	0	2026-05-22	2026-05-22	1	0	2060	2060	No	No	Si	Si	No	No	0	0e6c7928578d
10 - script PowerShell embebido	10	script PowerShell embebido	634	1	1	1	0	2026-05-22	2026-05-22	0	1	1893	1893	No	No	Si	Si	No	No	0	66f4adba34ef
10 - script PowerShell embebido	10	script PowerShell embebido	636	1	1	1	0	2026-05-22	2026-05-22	1	0	2099	2099	No	No	Si	Si	No	No	0	44c67e4617ff
10 - script PowerShell embebido	10	script PowerShell embebido	637	1	1	1	0	2026-05-21	2026-05-21	0	1	285	285	No	No	No	Si	Si	No	0	d22f6281fecf
10 - script PowerShell embebido	10	script PowerShell embebido	638	1	1	1	0	2026-05-21	2026-05-21	0	1	1897	1897	No	No	Si	Si	No	No	0	f2f25e31ac6f
10 - script PowerShell embebido	10	script PowerShell embebido	639	1	1	1	0	2026-05-21	2026-05-21	0	1	2040	2040	No	No	Si	Si	No	No	0	28a68c47757f
10 - script PowerShell embebido	10	script PowerShell embebido	640	1	1	1	0	2026-05-21	2026-05-21	0	1	2049	2049	No	No	Si	Si	No	No	0	0e65fb3a4854
10 - script PowerShell embebido	10	script PowerShell embebido	641	1	1	1	0	2026-05-21	2026-05-21	0	1	2065	2065	No	No	Si	Si	No	No	0	3a02f4279546
10 - script PowerShell embebido	10	script PowerShell embebido	642	1	1	1	0	2026-05-20	2026-05-20	0	1	1918	1918	No	No	Si	Si	No	No	0	2e1135771c9f
10 - script PowerShell embebido	10	script PowerShell embebido	643	1	1	1	0	2026-05-20	2026-05-20	0	1	4211	4211	No	No	No	Si	No	No	0	0a1396bd1f9f
10 - script PowerShell embebido	10	script PowerShell embebido	644	1	1	1	0	2026-05-20	2026-05-20	1	0	1918	1918	No	No	Si	Si	No	No	0	c5f4d4da58b7
10 - script PowerShell embebido	10	script PowerShell embebido	645	1	1	1	0	2026-05-20	2026-05-20	0	1	1906	1906	No	No	Si	Si	No	No	0	ddb0322e8755
10 - script PowerShell embebido	10	script PowerShell embebido	646	1	1	1	0	2026-05-20	2026-05-20	1	0	2065	2065	No	No	Si	Si	No	No	0	7f96a4545c28
10 - script PowerShell embebido	10	script PowerShell embebido	648	1	1	1	0	2026-05-20	2026-05-20	0	1	2056	2056	No	No	Si	Si	No	No	0	7533136e46d6
10 - script PowerShell embebido	10	script PowerShell embebido	649	1	1	1	0	2026-05-20	2026-05-20	1	0	2065	2065	No	No	Si	Si	No	No	0	c7874d132940
10 - script PowerShell embebido	10	script PowerShell embebido	650	1	1	1	0	2026-05-20	2026-05-20	0	1	4244	4244	No	No	No	Si	No	No	0	5985eee0b96b
10 - script PowerShell embebido	10	script PowerShell embebido	651	1	1	1	0	2026-05-20	2026-05-20	1	0	2093	2093	No	No	Si	Si	No	No	0	ea7698f3434d
10 - script PowerShell embebido	10	script PowerShell embebido	652	1	1	1	0	2026-05-20	2026-05-20	1	0	2030	2030	No	No	Si	Si	No	No	0	3ad0bb7dca00
10 - script PowerShell embebido	10	script PowerShell embebido	653	1	1	1	0	2026-05-20	2026-05-20	1	0	2032	2032	No	No	Si	Si	No	No	0	e0f0171b1f31
10 - script PowerShell embebido	10	script PowerShell embebido	654	1	1	1	0	2026-05-20	2026-05-20	1	0	2096	2096	No	No	Si	Si	No	No	0	acfa28592af6
10 - script PowerShell embebido	10	script PowerShell embebido	655	1	1	1	0	2026-05-20	2026-05-20	0	1	1861	1861	No	No	Si	Si	No	No	0	fd902dd14a30
10 - script PowerShell embebido	10	script PowerShell embebido	657	1	1	1	0	2026-05-20	2026-05-20	0	1	4214	4214	No	No	No	Si	No	No	0	d3331c44f760
10 - script PowerShell embebido	10	script PowerShell embebido	658	1	1	1	0	2026-05-19	2026-05-19	0	1	4219	4219	No	No	No	Si	No	No	0	7aa9fea6a9f0
10 - script PowerShell embebido	10	script PowerShell embebido	661	1	1	1	0	2026-05-18	2026-05-18	0	1	2036	2036	No	No	Si	Si	No	No	0	44d981574263
10 - script PowerShell embebido	10	script PowerShell embebido	662	1	1	1	0	2026-05-18	2026-05-18	0	1	2050	2050	No	No	Si	Si	No	No	0	ff77ee3c5fb7
10 - script PowerShell embebido	10	script PowerShell embebido	664	1	1	1	0	2026-05-17	2026-05-17	1	0	2094	2094	No	No	Si	Si	No	No	0	ba1a79c82972
10 - script PowerShell embebido	10	script PowerShell embebido	665	1	1	1	0	2026-05-17	2026-05-17	1	0	2044	2044	No	No	Si	Si	No	No	0	759b1ffc0ff9
10 - script PowerShell embebido	10	script PowerShell embebido	666	1	1	1	0	2026-05-17	2026-05-17	0	1	2086	2086	No	No	Si	Si	No	No	0	5692c27579ef
10 - script PowerShell embebido	10	script PowerShell embebido	667	1	1	1	0	2026-05-17	2026-05-17	1	0	2120	2120	No	No	Si	Si	No	No	0	b5c7b53a8e69
10 - script PowerShell embebido	10	script PowerShell embebido	668	1	1	1	0	2026-05-17	2026-05-17	1	0	2084	2084	No	No	Si	Si	No	No	0	1352876a87da
10 - script PowerShell embebido	10	script PowerShell embebido	669	1	1	1	0	2026-05-17	2026-05-17	0	1	2076	2076	No	No	Si	Si	No	No	0	db5cda477b2b
10 - script PowerShell embebido	10	script PowerShell embebido	670	1	1	1	0	2026-05-17	2026-05-17	1	0	2100	2100	No	No	Si	Si	No	No	0	6b9cc59bf73c
10 - script PowerShell embebido	10	script PowerShell embebido	671	1	1	1	0	2026-05-17	2026-05-17	0	1	1916	1916	No	No	Si	Si	No	No	0	83ab3e2f7a85
10 - script PowerShell embebido	10	script PowerShell embebido	672	1	1	1	0	2026-05-17	2026-05-17	1	0	2126	2126	No	No	Si	Si	No	No	0	2219d19a6f3b
10 - script PowerShell embebido	10	script PowerShell embebido	673	1	1	1	0	2026-05-17	2026-05-17	0	1	2098	2098	No	No	Si	Si	No	No	0	1abe5b6d51cf
10 - script PowerShell embebido	10	script PowerShell embebido	674	1	1	1	0	2026-05-17	2026-05-17	0	1	2043	2043	No	No	Si	Si	No	No	0	ff08a18586a9
10 - script PowerShell embebido	10	script PowerShell embebido	675	1	1	1	0	2026-05-17	2026-05-17	1	0	2107	2107	No	No	Si	Si	No	No	0	fb23481e1d1a
10 - script PowerShell embebido	10	script PowerShell embebido	676	1	1	1	0	2026-05-17	2026-05-17	0	1	1893	1893	No	No	Si	Si	No	No	0	0fda7caa5299
10 - script PowerShell embebido	10	script PowerShell embebido	677	1	1	1	0	2026-05-17	2026-05-17	0	1	2103	2103	No	No	Si	Si	No	No	0	c7f45a19bba7
10 - script PowerShell embebido	10	script PowerShell embebido	678	1	1	1	0	2026-05-17	2026-05-17	0	1	2099	2099	No	No	Si	Si	No	No	0	5da454054003
10 - script PowerShell embebido	10	script PowerShell embebido	679	1	1	1	0	2026-05-17	2026-05-17	0	1	1919	1919	No	No	Si	Si	No	No	0	511c6876673e
10 - script PowerShell embebido	10	script PowerShell embebido	680	1	1	1	0	2026-05-17	2026-05-17	1	0	2040	2040	No	No	Si	Si	No	No	0	55f7abab9115
10 - script PowerShell embebido	10	script PowerShell embebido	681	1	1	1	0	2026-05-17	2026-05-17	0	1	2040	2040	No	No	Si	Si	No	No	0	3c40c6aca861
10 - script PowerShell embebido	10	script PowerShell embebido	682	1	1	1	0	2026-05-16	2026-05-16	0	1	1140	1140	No	No	Si	Si	No	No	0	e16b38304290
10 - script PowerShell embebido	10	script PowerShell embebido	683	1	1	1	0	2026-05-15	2026-05-15	0	1	1670	1670	No	No	Si	Si	No	No	0	d0d467088b33
10 - script PowerShell embebido	10	script PowerShell embebido	684	1	1	1	0	2026-05-14	2026-05-14	1	0	1770	1770	No	No	Si	Si	No	No	0	5cca824e8ab8
10 - script PowerShell embebido	10	script PowerShell embebido	685	1	1	1	0	2026-05-14	2026-05-14	1	0	1113	1113	No	No	Si	Si	No	No	0	fcde113c9d28
10 - script PowerShell embebido	10	script PowerShell embebido	686	1	1	1	0	2026-05-14	2026-05-14	0	1	1679	1679	No	No	Si	Si	No	No	0	bdd14113cb9e
10 - script PowerShell embebido	10	script PowerShell embebido	687	1	1	1	0	2026-05-14	2026-05-14	1	0	1677	1677	No	No	Si	Si	No	No	0	0a6543fb4817
10 - script PowerShell embebido	10	script PowerShell embebido	688	1	1	1	0	2026-05-14	2026-05-14	1	0	1686	1686	No	No	Si	Si	No	No	0	87996f5b2076
10 - script PowerShell embebido	10	script PowerShell embebido	689	1	1	1	0	2026-05-14	2026-05-14	0	1	1651	1651	No	No	Si	Si	No	No	0	51025ed2696e
10 - script PowerShell embebido	10	script PowerShell embebido	690	1	1	1	0	2026-05-14	2026-05-14	1	0	1749	1749	No	No	Si	Si	No	No	0	ba6831f1f36a
10 - script PowerShell embebido	10	script PowerShell embebido	691	1	1	1	0	2026-05-14	2026-05-14	0	1	1119	1119	No	No	Si	Si	No	No	0	e4599d37108f
10 - script PowerShell embebido	10	script PowerShell embebido	692	1	1	1	0	2026-05-14	2026-05-14	0	1	1758	1758	No	No	Si	Si	No	No	0	ef931810cd07
10 - script PowerShell embebido	10	script PowerShell embebido	693	1	1	1	0	2026-05-14	2026-05-14	0	1	1796	1796	No	No	Si	Si	No	No	0	5cf666ccee82
10 - script PowerShell embebido	10	script PowerShell embebido	694	1	1	1	0	2026-05-14	2026-05-14	0	1	1686	1686	No	No	Si	Si	No	No	0	d484f295cbb3
10 - script PowerShell embebido	10	script PowerShell embebido	695	1	1	1	0	2026-05-14	2026-05-14	0	1	1692	1692	No	No	Si	Si	No	No	0	c743c7d84935
10 - script PowerShell embebido	10	script PowerShell embebido	696	1	1	1	0	2026-05-14	2026-05-14	0	1	1119	1119	No	No	Si	Si	No	No	0	e11e265dfca5
10 - script PowerShell embebido	10	script PowerShell embebido	697	1	1	1	0	2026-05-13	2026-05-13	0	1	4208	4208	No	No	No	Si	No	No	0	0a0abea6bb79
10 - script PowerShell embebido	10	script PowerShell embebido	698	1	1	1	0	2026-05-13	2026-05-13	0	1	4098	4098	No	No	No	Si	No	No	0	520643f7c3ca
10 - script PowerShell embebido	10	script PowerShell embebido	699	1	1	1	0	2026-05-13	2026-05-13	1	0	890	890	No	Si	No	No	Si	No	0	320060ee2190
10 - script PowerShell embebido	10	script PowerShell embebido	701	1	1	1	0	2026-05-13	2026-05-13	0	1	4273	4273	No	No	No	Si	No	No	0	14df3c42c807
10 - script PowerShell embebido	10	script PowerShell embebido	702	1	1	1	0	2026-05-13	2026-05-13	0	1	1764	1764	No	No	Si	Si	No	No	0	76ab95c2d8f7
10 - script PowerShell embebido	10	script PowerShell embebido	703	1	1	1	0	2026-05-13	2026-05-13	0	1	4215	4215	No	No	No	Si	No	No	0	e548ad649ff7
10 - script PowerShell embebido	10	script PowerShell embebido	704	1	1	1	0	2026-05-12	2026-05-12	0	1	4212	4212	No	No	No	Si	No	No	0	3ed7d01b6c53
10 - script PowerShell embebido	10	script PowerShell embebido	705	1	1	1	0	2026-05-12	2026-05-12	0	1	4224	4224	No	No	No	Si	No	No	0	2161c8c457bd
10 - script PowerShell embebido	10	script PowerShell embebido	706	1	1	1	0	2026-05-12	2026-05-12	0	1	4215	4215	No	No	No	Si	No	No	0	e7678f76adcd
10 - script PowerShell embebido	10	script PowerShell embebido	707	1	1	1	0	2026-05-12	2026-05-12	0	1	4246	4246	No	No	No	Si	No	No	0	1fb3c7b1f98a
10 - script PowerShell embebido	10	script PowerShell embebido	708	1	1	1	0	2026-05-12	2026-05-12	0	1	4235	4235	No	No	No	Si	No	No	0	ed1ec37e1b74
10 - script PowerShell embebido	10	script PowerShell embebido	709	1	1	1	0	2026-05-12	2026-05-12	0	1	4235	4235	No	No	No	Si	No	No	0	518ad9c367c8
10 - script PowerShell embebido	10	script PowerShell embebido	710	1	1	1	0	2026-05-12	2026-05-12	0	1	4221	4221	No	No	No	Si	No	No	0	6c8995ce68e0
10 - script PowerShell embebido	10	script PowerShell embebido	711	1	1	1	0	2026-05-12	2026-05-12	0	1	4234	4234	No	No	No	Si	No	No	0	43a5eabfc629
10 - script PowerShell embebido	10	script PowerShell embebido	712	1	1	1	0	2026-05-12	2026-05-12	0	1	4254	4254	No	No	No	Si	No	No	0	532cea6848ae
10 - script PowerShell embebido	10	script PowerShell embebido	713	1	1	1	0	2026-05-12	2026-05-12	0	1	4239	4239	No	No	No	Si	No	No	0	969362724273
10 - script PowerShell embebido	10	script PowerShell embebido	714	1	1	1	0	2026-05-12	2026-05-12	0	1	4245	4245	No	No	No	Si	No	No	0	668038c8a7bb
10 - script PowerShell embebido	10	script PowerShell embebido	715	1	1	1	0	2026-05-12	2026-05-12	0	1	4244	4244	No	No	No	Si	No	No	0	a148cb341d1f
10 - script PowerShell embebido	10	script PowerShell embebido	716	1	1	1	0	2026-05-12	2026-05-12	0	1	4231	4231	No	No	No	Si	No	No	0	94699e6f9340
10 - script PowerShell embebido	10	script PowerShell embebido	717	1	1	1	0	2026-05-12	2026-05-12	0	1	4238	4238	No	No	No	Si	No	No	0	034a8cab7d66
10 - script PowerShell embebido	10	script PowerShell embebido	718	1	1	1	0	2026-05-12	2026-05-12	0	1	4221	4221	No	No	No	Si	No	No	0	4568226edbed
10 - script PowerShell embebido	10	script PowerShell embebido	719	1	1	1	0	2026-05-12	2026-05-12	0	1	4225	4225	No	No	No	Si	No	No	0	2d08932cacf9
10 - script PowerShell embebido	10	script PowerShell embebido	720	1	1	1	0	2026-05-12	2026-05-12	0	1	4232	4232	No	No	No	Si	No	No	0	4076e49b9680
10 - script PowerShell embebido	10	script PowerShell embebido	721	1	1	1	0	2026-05-12	2026-05-12	0	1	4240	4240	No	No	No	Si	No	No	0	fee06fc80809
10 - script PowerShell embebido	10	script PowerShell embebido	722	1	1	1	0	2026-05-12	2026-05-12	0	1	4218	4218	No	No	No	Si	No	No	0	1b51d00d953d
10 - script PowerShell embebido	10	script PowerShell embebido	723	1	1	1	0	2026-05-12	2026-05-12	0	1	5871	5871	No	No	Si	Si	No	No	0	f3627cbc2088
10 - script PowerShell embebido	10	script PowerShell embebido	724	1	1	1	0	2026-05-12	2026-05-12	0	1	4253	4253	No	No	No	Si	No	No	0	05cb5bc98237
10 - script PowerShell embebido	10	script PowerShell embebido	725	1	1	1	0	2026-05-12	2026-05-12	0	1	4235	4235	No	No	No	Si	No	No	0	b6f628607280
10 - script PowerShell embebido	10	script PowerShell embebido	726	1	1	1	0	2026-05-12	2026-05-12	0	1	4235	4235	No	No	No	Si	No	No	0	36d09d3f58ac
10 - script PowerShell embebido	10	script PowerShell embebido	727	1	1	1	0	2026-05-12	2026-05-12	0	1	4244	4244	No	No	No	Si	No	No	0	dcc59388cf5c
10 - script PowerShell embebido	10	script PowerShell embebido	728	1	1	1	0	2026-05-12	2026-05-12	0	1	4228	4228	No	No	No	Si	No	No	0	7266535c22bd
10 - script PowerShell embebido	10	script PowerShell embebido	729	1	1	1	0	2026-05-12	2026-05-12	0	1	4265	4265	No	No	No	Si	No	No	0	8dd5044f7ed6
10 - script PowerShell embebido	10	script PowerShell embebido	730	1	1	1	0	2026-05-12	2026-05-12	0	1	4234	4234	No	No	No	Si	No	No	0	70aca914d5e2
10 - script PowerShell embebido	10	script PowerShell embebido	731	1	1	1	0	2026-05-12	2026-05-12	0	1	4237	4237	No	No	No	Si	No	No	0	06ab858ff391
10 - script PowerShell embebido	10	script PowerShell embebido	732	1	1	1	0	2026-05-12	2026-05-12	0	1	4212	4212	No	No	No	Si	No	No	0	005f13d24653
10 - script PowerShell embebido	10	script PowerShell embebido	733	1	1	1	0	2026-05-12	2026-05-12	0	1	4242	4242	No	No	No	Si	No	No	0	00ba1d09860f
10 - script PowerShell embebido	10	script PowerShell embebido	734	1	1	1	0	2026-05-12	2026-05-12	1	0	5269	5269	No	No	Si	Si	No	No	0	1d25343c4f9a
10 - script PowerShell embebido	10	script PowerShell embebido	735	1	1	1	0	2026-05-12	2026-05-12	0	1	4238	4238	No	No	No	Si	No	No	0	a54f8ad7f2eb
10 - script PowerShell embebido	10	script PowerShell embebido	736	1	1	1	0	2026-05-12	2026-05-12	0	1	4256	4256	No	No	No	Si	No	No	0	624ee1c849fe
10 - script PowerShell embebido	10	script PowerShell embebido	737	1	1	1	0	2026-05-12	2026-05-12	0	1	4235	4235	No	No	No	Si	No	No	0	8b3c68476dbb
10 - script PowerShell embebido	10	script PowerShell embebido	738	1	1	1	0	2026-05-12	2026-05-12	0	1	4243	4243	No	No	No	Si	No	No	0	63884e190fd0
10 - script PowerShell embebido	10	script PowerShell embebido	739	1	1	1	0	2026-05-12	2026-05-12	0	1	4215	4215	No	No	No	Si	No	No	0	f5fbdcfbd58c
10 - script PowerShell embebido	10	script PowerShell embebido	741	1	1	1	0	2026-05-12	2026-05-12	0	1	4266	4266	No	No	No	Si	No	No	0	d9628dac3bb6
10 - script PowerShell embebido	10	script PowerShell embebido	742	1	1	1	0	2026-05-12	2026-05-12	0	1	4234	4234	No	No	No	Si	No	No	0	f25f16d66c9a
10 - script PowerShell embebido	10	script PowerShell embebido	743	1	1	1	0	2026-05-12	2026-05-12	0	1	4233	4233	No	No	No	Si	No	No	0	3e37f8852e6e
10 - script PowerShell embebido	10	script PowerShell embebido	744	1	1	1	0	2026-05-12	2026-05-12	0	1	4237	4237	No	No	No	Si	No	No	0	9fcf7f3bf24c
10 - script PowerShell embebido	10	script PowerShell embebido	745	1	1	1	0	2026-05-12	2026-05-12	0	1	4240	4240	No	No	No	Si	No	No	0	705f6536b4c6
10 - script PowerShell embebido	10	script PowerShell embebido	746	1	1	1	0	2026-05-12	2026-05-12	0	1	4230	4230	No	No	No	Si	No	No	0	31386f5d02b0
10 - script PowerShell embebido	10	script PowerShell embebido	747	1	1	1	0	2026-05-12	2026-05-12	0	1	4250	4250	No	No	No	Si	No	No	0	ed982d582467
10 - script PowerShell embebido	10	script PowerShell embebido	748	1	1	1	0	2026-05-12	2026-05-12	0	1	1106	1106	No	No	Si	Si	No	No	0	5a50315b3688
10 - script PowerShell embebido	10	script PowerShell embebido	749	1	1	1	0	2026-05-12	2026-05-12	0	1	4090	4090	No	No	No	Si	No	No	0	f31b35e74a79
10 - script PowerShell embebido	10	script PowerShell embebido	750	1	1	1	0	2026-05-12	2026-05-12	0	1	4106	4106	No	No	No	Si	No	No	0	b429229700f1
10 - script PowerShell embebido	10	script PowerShell embebido	752	1	1	1	0	2026-05-11	2026-05-11	0	1	1109	1109	No	No	Si	Si	No	No	0	82e0f8b4e74a
10 - script PowerShell embebido	10	script PowerShell embebido	753	1	1	1	0	2026-05-11	2026-05-11	1	0	1676	1676	No	No	Si	Si	No	No	0	d3f3d2653d3f
10 - script PowerShell embebido	10	script PowerShell embebido	755	1	1	1	0	2026-05-11	2026-05-11	0	1	1813	1813	No	No	Si	Si	No	No	0	a6c3ac0f132f
10 - script PowerShell embebido	10	script PowerShell embebido	756	1	1	1	0	2026-05-11	2026-05-11	1	0	1780	1780	No	No	Si	Si	No	No	0	d0761b04a1fa
10 - script PowerShell embebido	10	script PowerShell embebido	757	1	1	1	0	2026-05-11	2026-05-11	1	0	1785	1785	No	No	Si	Si	No	No	0	244e446bc984
10 - script PowerShell embebido	10	script PowerShell embebido	758	1	1	1	0	2026-05-11	2026-05-11	1	0	1696	1696	No	No	Si	Si	No	No	0	f88b54d0c7f8
10 - script PowerShell embebido	10	script PowerShell embebido	759	1	1	1	0	2026-05-11	2026-05-11	1	0	1769	1769	No	No	Si	Si	No	No	0	61ae631475d5
10 - script PowerShell embebido	10	script PowerShell embebido	760	1	1	1	0	2026-05-11	2026-05-11	0	1	4238	4238	No	No	No	Si	No	No	0	2be96e55c6b2
10 - script PowerShell embebido	10	script PowerShell embebido	761	1	1	1	0	2026-05-11	2026-05-11	0	1	2220	2220	No	No	No	Si	No	No	0	ea040874e4bb
10 - script PowerShell embebido	10	script PowerShell embebido	762	1	1	1	0	2026-05-11	2026-05-11	1	0	1117	1117	No	No	Si	Si	No	No	0	63cf13746177
10 - script PowerShell embebido	10	script PowerShell embebido	763	1	1	1	0	2026-05-11	2026-05-11	0	1	1818	1818	No	No	Si	Si	No	No	0	a4498231e403
10 - script PowerShell embebido	10	script PowerShell embebido	764	1	1	1	0	2026-05-11	2026-05-11	1	0	1770	1770	No	No	Si	Si	No	No	0	cfa202a1a98e
10 - script PowerShell embebido	10	script PowerShell embebido	765	1	1	1	0	2026-05-11	2026-05-11	1	0	1115	1115	No	No	Si	Si	No	No	0	2b4d691314be
10 - script PowerShell embebido	10	script PowerShell embebido	766	1	1	1	0	2026-05-11	2026-05-11	1	0	1753	1753	No	No	Si	Si	No	No	0	acbb962895fa
10 - script PowerShell embebido	10	script PowerShell embebido	767	1	1	1	0	2026-05-11	2026-05-11	0	1	1756	1756	No	No	Si	Si	No	No	0	dfdeffa136b1
10 - script PowerShell embebido	10	script PowerShell embebido	768	1	1	1	0	2026-05-11	2026-05-11	0	1	1792	1792	No	No	Si	Si	No	No	0	cc30a6567961
10 - script PowerShell embebido	10	script PowerShell embebido	769	1	1	1	0	2026-05-11	2026-05-11	0	1	1816	1816	No	No	Si	Si	No	No	0	03accda45b24
10 - script PowerShell embebido	10	script PowerShell embebido	770	1	1	1	0	2026-05-10	2026-05-10	1	0	1113	1113	No	No	Si	Si	No	No	0	e7f17345bd79
10 - script PowerShell embebido	10	script PowerShell embebido	771	1	1	1	0	2026-05-10	2026-05-10	1	0	1675	1675	No	No	Si	Si	No	No	0	2bb88b6d6ef6
10 - script PowerShell embebido	10	script PowerShell embebido	772	1	1	1	0	2026-05-10	2026-05-10	1	0	1805	1805	No	No	Si	Si	No	No	0	6dc2011d926c
10 - script PowerShell embebido	10	script PowerShell embebido	773	1	1	1	0	2026-05-10	2026-05-10	0	1	1121	1121	No	No	Si	Si	No	No	0	5d02d09d53e0
10 - script PowerShell embebido	10	script PowerShell embebido	774	1	1	1	0	2026-05-10	2026-05-10	0	1	1792	1792	No	No	Si	Si	No	No	0	945383cef7e5
10 - script PowerShell embebido	10	script PowerShell embebido	775	1	1	1	0	2026-05-10	2026-05-10	0	1	1689	1689	No	No	Si	Si	No	No	0	491dd1439bc7
10 - script PowerShell embebido	10	script PowerShell embebido	776	1	1	1	0	2026-05-10	2026-05-10	0	1	1116	1116	No	No	Si	Si	No	No	0	d0aebaf6036e
10 - script PowerShell embebido	10	script PowerShell embebido	777	1	1	1	0	2026-05-10	2026-05-10	1	0	1686	1686	No	No	Si	Si	No	No	0	e8c4da86e59e
10 - script PowerShell embebido	10	script PowerShell embebido	778	1	1	1	0	2026-05-10	2026-05-10	1	0	1118	1118	No	No	Si	Si	No	No	0	6402d8e0dfbd
10 - script PowerShell embebido	10	script PowerShell embebido	779	1	1	1	0	2026-05-10	2026-05-10	0	1	1792	1792	No	No	Si	Si	No	No	0	cecaf471c146
10 - script PowerShell embebido	10	script PowerShell embebido	780	1	1	1	0	2026-05-10	2026-05-10	0	1	1127	1127	No	No	Si	Si	No	No	0	b4ad721e8c4a
10 - script PowerShell embebido	10	script PowerShell embebido	781	1	1	1	0	2026-05-10	2026-05-10	1	0	1766	1766	No	No	Si	Si	No	No	0	a318de97c56d
10 - script PowerShell embebido	10	script PowerShell embebido	782	1	1	1	0	2026-05-10	2026-05-10	1	0	1691	1691	No	No	Si	Si	No	No	0	7daea51ff152
10 - script PowerShell embebido	10	script PowerShell embebido	783	1	1	1	0	2026-05-10	2026-05-10	0	1	1777	1777	No	No	Si	Si	No	No	0	6e0a407b9c0d
10 - script PowerShell embebido	10	script PowerShell embebido	784	1	1	1	0	2026-05-10	2026-05-10	1	0	1120	1120	No	No	Si	Si	No	No	0	a720daa181fc
10 - script PowerShell embebido	10	script PowerShell embebido	785	1	1	1	0	2026-05-10	2026-05-10	0	1	1111	1111	No	No	Si	Si	No	No	0	04ead84741cb
10 - script PowerShell embebido	10	script PowerShell embebido	786	1	1	1	0	2026-05-10	2026-05-10	1	0	1122	1122	No	No	Si	Si	No	No	0	98111d333b9a
10 - script PowerShell embebido	10	script PowerShell embebido	787	1	1	1	0	2026-05-10	2026-05-10	0	1	1684	1684	No	No	Si	Si	No	No	0	5088258b7cda
10 - script PowerShell embebido	10	script PowerShell embebido	788	1	1	1	0	2026-05-10	2026-05-10	0	1	1802	1802	No	No	Si	Si	No	No	0	1afe76d90cad
10 - script PowerShell embebido	10	script PowerShell embebido	789	1	1	1	0	2026-05-10	2026-05-10	0	1	1122	1122	No	No	Si	Si	No	No	0	6646473a2998
10 - script PowerShell embebido	10	script PowerShell embebido	790	1	1	1	0	2026-05-10	2026-05-10	1	0	1787	1787	No	No	Si	Si	No	No	0	20b7b5c8e359
10 - script PowerShell embebido	10	script PowerShell embebido	791	1	1	1	0	2026-05-10	2026-05-10	1	0	1806	1806	No	No	Si	Si	No	No	0	6056c32d8cf4
10 - script PowerShell embebido	10	script PowerShell embebido	792	1	1	1	0	2026-05-10	2026-05-10	0	1	1802	1802	No	No	Si	Si	No	No	0	9f9c2bfb857d
10 - script PowerShell embebido	10	script PowerShell embebido	793	1	1	1	0	2026-05-10	2026-05-10	0	1	1128	1128	No	No	Si	Si	No	No	0	c51d41bf04b2
10 - script PowerShell embebido	10	script PowerShell embebido	794	1	1	1	0	2026-05-10	2026-05-10	1	0	1761	1761	No	No	Si	Si	No	No	0	05986da02e3b
10 - script PowerShell embebido	10	script PowerShell embebido	795	1	1	1	0	2026-05-10	2026-05-10	0	1	1785	1785	No	No	Si	Si	No	No	0	c1bfea1226f2
10 - script PowerShell embebido	10	script PowerShell embebido	796	1	1	1	0	2026-05-10	2026-05-10	0	1	1694	1694	No	No	Si	Si	No	No	0	8b78b4030d26
10 - script PowerShell embebido	10	script PowerShell embebido	797	1	1	1	0	2026-05-10	2026-05-10	0	1	1758	1758	No	No	Si	Si	No	No	0	9d8843ad9971
10 - script PowerShell embebido	10	script PowerShell embebido	798	1	1	1	0	2026-05-10	2026-05-10	1	0	1747	1747	No	No	Si	Si	No	No	0	39ca3f8aa5c7
10 - script PowerShell embebido	10	script PowerShell embebido	799	1	1	1	0	2026-05-10	2026-05-10	0	1	1754	1754	No	No	Si	Si	No	No	0	020a736027a0
10 - script PowerShell embebido	10	script PowerShell embebido	800	1	1	1	0	2026-05-10	2026-05-10	1	0	1121	1121	No	No	Si	Si	No	No	0	6c23c772117b
10 - script PowerShell embebido	10	script PowerShell embebido	801	1	1	1	0	2026-05-10	2026-05-10	1	0	1778	1778	No	No	Si	Si	No	No	0	3b0c97d8e6e3
10 - script PowerShell embebido	10	script PowerShell embebido	802	1	1	1	0	2026-05-10	2026-05-10	0	1	1128	1128	No	No	Si	Si	No	No	0	426dfa643c29
10 - script PowerShell embebido	10	script PowerShell embebido	803	1	1	1	0	2026-05-10	2026-05-10	1	0	1800	1800	No	No	Si	Si	No	No	0	b1006309bb82
10 - script PowerShell embebido	10	script PowerShell embebido	804	1	1	1	0	2026-05-10	2026-05-10	1	0	1668	1668	No	No	Si	Si	No	No	0	0832edbf36e7
10 - script PowerShell embebido	10	script PowerShell embebido	805	1	1	1	0	2026-05-10	2026-05-10	0	1	1129	1129	No	No	Si	Si	No	No	0	c53e1687ea3d
10 - script PowerShell embebido	10	script PowerShell embebido	806	1	1	1	0	2026-05-10	2026-05-10	0	1	1693	1693	No	No	Si	Si	No	No	0	ca65990a08d9
10 - script PowerShell embebido	10	script PowerShell embebido	807	1	1	1	0	2026-05-10	2026-05-10	0	1	1796	1796	No	No	Si	Si	No	No	0	2d0bad1dc30a
10 - script PowerShell embebido	10	script PowerShell embebido	808	1	1	1	0	2026-05-10	2026-05-10	1	0	1131	1131	No	No	Si	Si	No	No	0	2144a8894d5e
10 - script PowerShell embebido	10	script PowerShell embebido	809	1	1	1	0	2026-05-10	2026-05-10	1	0	1672	1672	No	No	Si	Si	No	No	0	f384fc03a783
10 - script PowerShell embebido	10	script PowerShell embebido	810	1	1	1	0	2026-05-10	2026-05-10	1	0	1673	1673	No	No	Si	Si	No	No	0	92d42c1bb7c6
10 - script PowerShell embebido	10	script PowerShell embebido	811	1	1	1	0	2026-05-10	2026-05-10	0	1	1117	1117	No	No	Si	Si	No	No	0	d5220149df1b
10 - script PowerShell embebido	10	script PowerShell embebido	812	1	1	1	0	2026-05-10	2026-05-10	1	0	1662	1662	No	No	Si	Si	No	No	0	935600d1272f
10 - script PowerShell embebido	10	script PowerShell embebido	813	1	1	1	0	2026-05-10	2026-05-10	1	0	1133	1133	No	No	Si	Si	No	No	0	5edecc422630
10 - script PowerShell embebido	10	script PowerShell embebido	814	1	1	1	0	2026-05-10	2026-05-10	1	0	1773	1773	No	No	Si	Si	No	No	0	e1939d96257c
10 - script PowerShell embebido	10	script PowerShell embebido	815	1	1	1	0	2026-05-10	2026-05-10	0	1	1758	1758	No	No	Si	Si	No	No	0	de91406878ca
10 - script PowerShell embebido	10	script PowerShell embebido	816	1	1	1	0	2026-05-10	2026-05-10	0	1	1116	1116	No	No	Si	Si	No	No	0	44bb256d6fa7
10 - script PowerShell embebido	10	script PowerShell embebido	817	1	1	1	0	2026-05-10	2026-05-10	1	0	4095	4095	No	No	No	Si	No	No	0	4069c181aabe
10 - script PowerShell embebido	10	script PowerShell embebido	818	1	1	1	0	2026-05-10	2026-05-10	1	0	4088	4088	No	No	No	Si	No	No	0	1bf4eabff868
10 - script PowerShell embebido	10	script PowerShell embebido	819	1	1	1	0	2026-05-10	2026-05-10	0	1	4087	4087	No	No	No	Si	No	No	0	1043c9f40a7a
10 - script PowerShell embebido	10	script PowerShell embebido	820	1	1	1	0	2026-05-10	2026-05-10	0	1	4189	4189	No	No	No	Si	No	No	0	a8d640308ac6
10 - script PowerShell embebido	10	script PowerShell embebido	821	1	1	1	0	2026-05-10	2026-05-10	1	0	1805	1805	No	No	Si	Si	No	No	0	0352f625fc66
10 - script PowerShell embebido	10	script PowerShell embebido	822	1	1	1	0	2026-05-10	2026-05-10	1	0	1780	1780	No	No	Si	Si	No	No	0	745a04605644
10 - script PowerShell embebido	10	script PowerShell embebido	823	1	1	1	0	2026-05-10	2026-05-10	1	0	1798	1798	No	No	Si	Si	No	No	0	8cac48689d7a
10 - script PowerShell embebido	10	script PowerShell embebido	824	1	1	1	0	2026-05-10	2026-05-10	0	1	1124	1124	No	No	Si	Si	No	No	0	745038256c74
10 - script PowerShell embebido	10	script PowerShell embebido	825	1	1	1	0	2026-05-10	2026-05-10	1	0	1123	1123	No	No	Si	Si	No	No	0	a9e5d227dc87
10 - script PowerShell embebido	10	script PowerShell embebido	826	1	1	1	0	2026-05-10	2026-05-10	1	0	1682	1682	No	No	Si	Si	No	No	0	a47210ccb3ae
10 - script PowerShell embebido	10	script PowerShell embebido	827	1	1	1	0	2026-05-10	2026-05-10	0	1	1121	1121	No	No	Si	Si	No	No	0	81df2c33f02e
10 - script PowerShell embebido	10	script PowerShell embebido	828	1	1	1	0	2026-05-10	2026-05-10	0	1	1125	1125	No	No	Si	Si	No	No	0	36bcf5fe4b22
10 - script PowerShell embebido	10	script PowerShell embebido	829	1	1	1	0	2026-05-10	2026-05-10	1	0	1115	1115	No	No	Si	Si	No	No	0	6b28c10b00ba
10 - script PowerShell embebido	10	script PowerShell embebido	830	1	1	1	0	2026-05-10	2026-05-10	1	0	1781	1781	No	No	Si	Si	No	No	0	9eedf0eadaf6
10 - script PowerShell embebido	10	script PowerShell embebido	831	1	1	1	0	2026-05-10	2026-05-10	0	1	1125	1125	No	No	Si	Si	No	No	0	fe503db88f9b
10 - script PowerShell embebido	10	script PowerShell embebido	832	1	1	1	0	2026-05-10	2026-05-10	1	0	1757	1757	No	No	Si	Si	No	No	0	87ea7653f65a
10 - script PowerShell embebido	10	script PowerShell embebido	833	1	1	1	0	2026-05-10	2026-05-10	0	1	1671	1671	No	No	Si	Si	No	No	0	5318cd990733
10 - script PowerShell embebido	10	script PowerShell embebido	834	1	1	1	0	2026-05-10	2026-05-10	0	1	1695	1695	No	No	Si	Si	No	No	0	f22e677ca104
10 - script PowerShell embebido	10	script PowerShell embebido	835	1	1	1	0	2026-05-10	2026-05-10	0	1	1695	1695	No	No	Si	Si	No	No	0	bb0536f3157a
10 - script PowerShell embebido	10	script PowerShell embebido	836	1	1	1	0	2026-05-10	2026-05-10	1	0	1697	1697	No	No	Si	Si	No	No	0	cbedf28e12c6
10 - script PowerShell embebido	10	script PowerShell embebido	837	1	1	1	0	2026-05-10	2026-05-10	0	1	4099	4099	No	No	No	Si	No	No	0	a87a3b76c419
10 - script PowerShell embebido	10	script PowerShell embebido	838	1	1	1	0	2026-05-10	2026-05-10	0	1	4244	4244	No	No	No	Si	No	No	0	b6be3f2bfa80
10 - script PowerShell embebido	10	script PowerShell embebido	839	1	1	1	0	2026-05-10	2026-05-10	0	1	4101	4101	No	No	No	Si	No	No	0	699360c84ea7
10 - script PowerShell embebido	10	script PowerShell embebido	840	1	1	1	0	2026-05-10	2026-05-10	0	1	4091	4091	No	No	No	Si	No	No	0	6d468ecb1b5b
10 - script PowerShell embebido	10	script PowerShell embebido	841	1	1	1	0	2026-05-10	2026-05-10	0	1	2109	2109	No	No	No	Si	No	No	0	7b052f1a785d
10 - script PowerShell embebido	10	script PowerShell embebido	842	1	1	1	0	2026-05-10	2026-05-10	0	1	4253	4253	No	No	No	Si	No	No	0	54693609f19d
10 - script PowerShell embebido	10	script PowerShell embebido	843	1	1	1	0	2026-05-10	2026-05-10	0	1	4258	4258	No	No	No	Si	No	No	0	79df5051705e
10 - script PowerShell embebido	10	script PowerShell embebido	844	1	1	1	0	2026-05-10	2026-05-10	0	1	4233	4233	No	No	No	Si	No	No	0	65bf96956e88
10 - script PowerShell embebido	10	script PowerShell embebido	845	1	1	1	0	2026-05-10	2026-05-10	0	1	2220	2220	No	No	No	Si	No	No	0	bacad4e438eb
10 - script PowerShell embebido	10	script PowerShell embebido	846	1	1	1	0	2026-05-10	2026-05-10	0	1	4241	4241	No	No	No	Si	No	No	0	a6240bdb20b0
10 - script PowerShell embebido	10	script PowerShell embebido	847	1	1	1	0	2026-05-10	2026-05-10	1	0	4093	4093	No	No	No	Si	No	No	0	122f29a2ecf0
10 - script PowerShell embebido	10	script PowerShell embebido	848	1	1	1	0	2026-05-10	2026-05-10	0	1	4147	4147	No	No	No	Si	No	No	0	2bf6ac5bfc95
10 - script PowerShell embebido	10	script PowerShell embebido	849	1	1	1	0	2026-05-10	2026-05-10	0	1	1786	1786	No	No	Si	Si	No	No	0	2e353e9ea894
10 - script PowerShell embebido	10	script PowerShell embebido	850	1	1	1	0	2026-05-09	2026-05-09	1	0	1799	1799	No	No	Si	Si	No	No	0	ee744983f9f7
10 - script PowerShell embebido	10	script PowerShell embebido	851	1	1	1	0	2026-05-09	2026-05-09	0	1	56	56	No	Si	No	Si	No	No	0	dc49d69025be
10 - script PowerShell embebido	10	script PowerShell embebido	852	1	1	1	0	2026-05-08	2026-05-08	0	1	1114	1114	No	No	Si	Si	No	No	0	2a521f320f35
10 - script PowerShell embebido	10	script PowerShell embebido	853	1	1	1	0	2026-05-08	2026-05-08	0	1	4079	4079	No	No	No	Si	No	No	0	df0f45a4008c
10 - script PowerShell embebido	10	script PowerShell embebido	854	1	1	1	0	2026-05-08	2026-05-08	0	1	1114	1114	No	No	Si	Si	No	No	0	d99135da9bae
10 - script PowerShell embebido	10	script PowerShell embebido	855	1	1	1	0	2026-05-08	2026-05-08	0	1	1114	1114	No	No	Si	Si	No	No	0	5523ad9c35c6
10 - script PowerShell embebido	10	script PowerShell embebido	856	1	1	1	0	2026-05-08	2026-05-08	0	1	1129	1129	No	No	Si	Si	No	No	0	0fa093f78051
10 - script PowerShell embebido	10	script PowerShell embebido	857	1	1	1	0	2026-05-08	2026-05-08	1	0	1707	1707	No	No	Si	Si	No	No	0	43c6b1781964
10 - script PowerShell embebido	10	script PowerShell embebido	858	1	1	1	0	2026-05-08	2026-05-08	0	1	4148	4148	No	No	No	Si	No	No	0	215e3bd4606e
10 - script PowerShell embebido	10	script PowerShell embebido	860	1	1	1	0	2026-05-08	2026-05-08	0	1	1739	1739	No	No	Si	Si	No	No	0	e3f594456577
10 - script PowerShell embebido	10	script PowerShell embebido	861	1	1	1	0	2026-05-08	2026-05-08	0	1	1728	1728	No	No	Si	Si	No	No	0	5b277c646ba2
10 - script PowerShell embebido	10	script PowerShell embebido	862	1	1	1	0	2026-05-07	2026-05-07	0	1	4245	4245	No	No	No	Si	No	No	0	c5bb0ce9dcf7
10 - script PowerShell embebido	10	script PowerShell embebido	863	1	1	1	0	2026-05-07	2026-05-07	0	1	4171	4171	No	No	No	Si	No	No	0	e6e92c67eb94
10 - script PowerShell embebido	10	script PowerShell embebido	864	1	1	1	0	2026-05-06	2026-05-06	1	0	1128	1128	No	No	Si	Si	No	No	0	ad49966676b3
10 - script PowerShell embebido	10	script PowerShell embebido	865	1	1	1	0	2026-05-06	2026-05-06	0	1	1768	1768	No	No	Si	Si	No	No	0	65f25383d820
10 - script PowerShell embebido	10	script PowerShell embebido	867	1	1	1	0	2026-05-06	2026-05-06	1	0	1651	1651	No	No	Si	Si	No	No	0	a1fd7199af26
10 - script PowerShell embebido	10	script PowerShell embebido	871	1	1	1	0	2026-05-06	2026-05-06	0	1	4247	4247	No	No	No	Si	No	No	0	f000ac8332a9
10 - script PowerShell embebido	10	script PowerShell embebido	872	1	1	1	0	2026-05-06	2026-05-06	0	1	4112	4112	No	No	No	Si	No	No	0	4611ef4b09b6
10 - script PowerShell embebido	10	script PowerShell embebido	873	1	1	1	0	2026-05-05	2026-05-05	1	0	1672	1672	No	No	Si	Si	No	No	0	2b46888be65c
10 - script PowerShell embebido	10	script PowerShell embebido	874	1	1	1	0	2026-05-05	2026-05-05	0	1	1115	1115	No	No	Si	Si	No	No	0	803489c9d994
10 - script PowerShell embebido	10	script PowerShell embebido	875	1	1	1	0	2026-05-05	2026-05-05	0	1	1683	1683	No	No	Si	Si	No	No	0	1619c1efa392
10 - script PowerShell embebido	10	script PowerShell embebido	876	1	1	1	0	2026-05-05	2026-05-05	0	1	4110	4110	No	No	No	Si	No	No	0	2d79498bfdcd
10 - script PowerShell embebido	10	script PowerShell embebido	877	1	1	1	0	2026-05-05	2026-05-05	1	0	1699	1699	No	No	Si	Si	No	No	0	a93aaaa70b89
10 - script PowerShell embebido	10	script PowerShell embebido	878	1	1	1	0	2026-05-05	2026-05-05	1	0	1758	1758	No	No	Si	Si	No	No	0	bcf858186f16
10 - script PowerShell embebido	10	script PowerShell embebido	879	1	1	1	0	2026-05-05	2026-05-05	1	0	1754	1754	No	No	Si	Si	No	No	0	626aae2b5d3f
10 - script PowerShell embebido	10	script PowerShell embebido	880	1	1	1	0	2026-05-05	2026-05-05	0	1	1759	1759	No	No	Si	Si	No	No	0	1a052ff2945e
10 - script PowerShell embebido	10	script PowerShell embebido	881	1	1	1	0	2026-05-05	2026-05-05	0	1	1115	1115	No	No	Si	Si	No	No	0	a5b856751573
10 - script PowerShell embebido	10	script PowerShell embebido	882	1	1	1	0	2026-05-05	2026-05-05	0	1	1687	1687	No	No	Si	Si	No	No	0	de9313e9edf8
10 - script PowerShell embebido	10	script PowerShell embebido	883	1	1	1	0	2026-05-05	2026-05-05	0	1	1734	1734	No	No	Si	Si	No	No	0	1d9ea186a96f
10 - script PowerShell embebido	10	script PowerShell embebido	884	1	1	1	0	2026-05-05	2026-05-05	0	1	4253	4253	No	No	No	Si	No	No	0	74358c0ea2b6
10 - script PowerShell embebido	10	script PowerShell embebido	885	1	1	1	0	2026-05-05	2026-05-05	1	0	1682	1682	No	No	Si	Si	No	No	0	0bd104d30e7f
10 - script PowerShell embebido	10	script PowerShell embebido	886	1	1	1	0	2026-05-05	2026-05-05	0	1	1105	1105	No	No	Si	Si	No	No	0	d64686e1954b
10 - script PowerShell embebido	10	script PowerShell embebido	887	1	1	1	0	2026-05-05	2026-05-05	0	1	2284	2284	No	No	Si	Si	No	No	0	0f7d56a8b4ff
10 - script PowerShell embebido	10	script PowerShell embebido	888	1	1	1	0	2026-05-04	2026-05-04	0	1	1115	1115	No	No	Si	Si	No	No	0	4f8d208c5b9a
10 - script PowerShell embebido	10	script PowerShell embebido	889	1	1	1	0	2026-05-04	2026-05-04	1	0	1687	1687	No	No	Si	Si	No	No	0	f7a48e62e48d
10 - script PowerShell embebido	10	script PowerShell embebido	890	1	1	1	0	2026-05-04	2026-05-04	0	1	4153	4153	No	No	No	Si	No	No	0	2431564e02d8
10 - script PowerShell embebido	10	script PowerShell embebido	892	1	1	1	0	2026-05-04	2026-05-04	0	1	1674	1674	No	No	Si	Si	No	No	0	3fc45f2d1ba8
10 - script PowerShell embebido	10	script PowerShell embebido	893	1	1	1	0	2026-05-03	2026-05-03	0	1	4106	4106	No	No	No	Si	No	No	0	495f5cad0d13
10 - script PowerShell embebido	10	script PowerShell embebido	894	1	1	1	0	2026-05-03	2026-05-03	0	1	4125	4125	No	No	No	Si	No	No	0	df041f1f2e0a
10 - script PowerShell embebido	10	script PowerShell embebido	897	1	1	1	0	2026-05-03	2026-05-03	0	1	4097	4097	No	No	No	Si	No	No	0	92d6b1435950
10 - script PowerShell embebido	10	script PowerShell embebido	898	1	1	1	0	2026-05-03	2026-05-03	0	1	1665	1665	No	No	Si	Si	No	No	0	c6289deed514
10 - script PowerShell embebido	10	script PowerShell embebido	899	1	1	1	0	2026-05-03	2026-05-03	0	1	4117	4117	No	No	No	Si	No	No	0	8767481ba6a6
10 - script PowerShell embebido	10	script PowerShell embebido	900	1	1	1	0	2026-05-02	2026-05-02	0	1	1782	1782	No	No	Si	Si	No	No	0	d3e7efb1f8de
10 - script PowerShell embebido	10	script PowerShell embebido	901	1	1	1	0	2026-05-02	2026-05-02	1	0	1683	1683	No	No	Si	Si	No	No	0	0ce3e1b9b8cb
10 - script PowerShell embebido	10	script PowerShell embebido	902	1	1	1	0	2026-05-02	2026-05-02	1	0	1131	1131	No	No	Si	Si	No	No	0	f6c23fcd8efd
10 - script PowerShell embebido	10	script PowerShell embebido	903	1	1	1	0	2026-05-02	2026-05-02	1	0	1787	1787	No	No	Si	Si	No	No	0	f59325461bc2
10 - script PowerShell embebido	10	script PowerShell embebido	904	1	1	1	0	2026-05-02	2026-05-02	1	0	1723	1723	No	No	Si	Si	No	No	0	91d1bf95ef3f
10 - script PowerShell embebido	10	script PowerShell embebido	905	1	1	1	0	2026-05-02	2026-05-02	0	1	4106	4106	No	No	No	Si	No	No	0	210f05ecd3c5
10 - script PowerShell embebido	10	script PowerShell embebido	906	1	1	1	0	2026-05-02	2026-05-02	0	1	1715	1715	No	No	Si	Si	No	No	0	54e012d3a6b9
10 - script PowerShell embebido	10	script PowerShell embebido	907	1	1	1	0	2026-05-02	2026-05-02	0	1	1724	1724	No	No	Si	Si	No	No	0	39bd524e8fc8
10 - script PowerShell embebido	10	script PowerShell embebido	908	1	1	1	0	2026-05-02	2026-05-02	1	0	1738	1738	No	No	Si	Si	No	No	0	9e39818e903b
10 - script PowerShell embebido	10	script PowerShell embebido	909	1	1	1	0	2026-05-02	2026-05-02	1	0	1748	1748	No	No	Si	Si	No	No	0	dcd5aeae4c3f
10 - script PowerShell embebido	10	script PowerShell embebido	910	1	1	1	0	2026-05-02	2026-05-02	1	0	1718	1718	No	No	Si	Si	No	No	0	a4e2126afb7e
10 - script PowerShell embebido	10	script PowerShell embebido	911	1	1	1	0	2026-05-02	2026-05-02	0	1	1748	1748	No	No	Si	Si	No	No	0	4e5185b009b6
10 - script PowerShell embebido	10	script PowerShell embebido	912	1	1	0	0	2026-05-02	2026-05-02	0	1	1096	1096	No	No	Si	Si	No	No	1	8c0c3fd6a406
10 - script PowerShell embebido	10	script PowerShell embebido	913	1	1	1	0	2026-05-02	2026-05-02	1	0	1636	1636	No	No	Si	Si	No	No	0	a04b967652f2
10 - script PowerShell embebido	10	script PowerShell embebido	914	1	1	1	0	2026-05-02	2026-05-02	0	1	1711	1711	No	No	Si	Si	No	No	0	970e22838ec6
10 - script PowerShell embebido	10	script PowerShell embebido	915	1	1	0	0	2026-05-02	2026-05-02	1	0	1730	1730	No	No	Si	Si	No	No	1	93bcb2fef891
10 - script PowerShell embebido	10	script PowerShell embebido	916	1	1	1	0	2026-05-02	2026-05-02	1	0	1746	1746	No	No	Si	Si	No	No	0	ea101ed08b26
10 - script PowerShell embebido	10	script PowerShell embebido	917	1	1	1	0	2026-05-02	2026-05-02	0	1	1702	1702	No	No	Si	Si	No	No	0	9745a77a6b93
10 - script PowerShell embebido	10	script PowerShell embebido	918	1	1	1	0	2026-05-02	2026-05-02	0	1	1686	1686	No	No	Si	Si	No	No	0	226461a9b010
10 - script PowerShell embebido	10	script PowerShell embebido	919	1	1	1	0	2026-05-02	2026-05-02	0	1	1720	1720	No	No	Si	Si	No	No	0	749bda0edabf
10 - script PowerShell embebido	10	script PowerShell embebido	920	1	1	1	0	2026-05-02	2026-05-02	0	1	1701	1701	No	No	Si	Si	No	No	0	12254a370e83
10 - script PowerShell embebido	10	script PowerShell embebido	921	1	1	1	0	2026-05-02	2026-05-02	0	1	1708	1708	No	No	Si	Si	No	No	0	3f4ced3f02dd
10 - script PowerShell embebido	10	script PowerShell embebido	922	1	1	1	0	2026-05-02	2026-05-02	0	1	1110	1110	No	No	Si	Si	No	No	0	93e8ca718788
10 - script PowerShell embebido	10	script PowerShell embebido	923	1	1	1	0	2026-05-02	2026-05-02	0	1	1755	1755	No	No	Si	Si	No	No	0	b18b08b69a2c
10 - script PowerShell embebido	10	script PowerShell embebido	924	1	1	1	0	2026-05-02	2026-05-02	0	1	1717	1717	No	No	Si	Si	No	No	0	1a470509efb0
10 - script PowerShell embebido	10	script PowerShell embebido	925	1	1	1	0	2026-05-02	2026-05-02	1	0	1106	1106	No	No	Si	Si	No	No	0	5aad1dbee94d
10 - script PowerShell embebido	10	script PowerShell embebido	926	1	1	1	0	2026-05-02	2026-05-02	1	0	1680	1680	No	No	Si	Si	No	No	0	bd06d9118711
10 - script PowerShell embebido	10	script PowerShell embebido	927	1	1	1	0	2026-05-02	2026-05-02	0	1	1108	1108	No	No	Si	Si	No	No	0	44dd9e2774d1
10 - script PowerShell embebido	10	script PowerShell embebido	928	1	1	1	0	2026-05-02	2026-05-02	0	1	1639	1639	No	No	Si	Si	No	No	0	403d90b9fa73
10 - script PowerShell embebido	10	script PowerShell embebido	929	1	1	1	0	2026-05-02	2026-05-02	0	1	1721	1721	No	No	Si	Si	No	No	0	5f62bff1e2ba
10 - script PowerShell embebido	10	script PowerShell embebido	930	1	1	1	0	2026-05-02	2026-05-02	0	1	1756	1756	No	No	Si	Si	No	No	0	6fde859280c2
10 - script PowerShell embebido	10	script PowerShell embebido	931	1	1	1	0	2026-05-02	2026-05-02	0	1	1117	1117	No	No	Si	Si	No	No	0	12652257a5eb
10 - script PowerShell embebido	10	script PowerShell embebido	932	1	1	1	0	2026-05-02	2026-05-02	1	0	1118	1118	No	No	Si	Si	No	No	0	9b2f8700c65f
10 - script PowerShell embebido	10	script PowerShell embebido	933	1	1	1	0	2026-05-02	2026-05-02	1	0	1748	1748	No	No	Si	Si	No	No	0	dfa8ff998a4d
10 - script PowerShell embebido	10	script PowerShell embebido	934	1	1	1	0	2026-05-02	2026-05-02	1	0	1099	1099	No	No	Si	Si	No	No	0	eb970077bf57
10 - script PowerShell embebido	10	script PowerShell embebido	935	1	1	1	0	2026-05-02	2026-05-02	0	1	1734	1734	No	No	Si	Si	No	No	0	88b81476fc06
10 - script PowerShell embebido	10	script PowerShell embebido	936	1	1	1	0	2026-05-02	2026-05-02	1	0	1648	1648	No	No	Si	Si	No	No	0	6a0b3e89eb4f
10 - script PowerShell embebido	10	script PowerShell embebido	937	1	1	1	0	2026-05-02	2026-05-02	1	0	1738	1738	No	No	Si	Si	No	No	0	31174be4a526
10 - script PowerShell embebido	10	script PowerShell embebido	938	1	1	1	0	2026-05-02	2026-05-02	1	0	1620	1620	No	No	Si	Si	No	No	0	2d9c7b41601b
10 - script PowerShell embebido	10	script PowerShell embebido	940	1	1	1	0	2026-05-02	2026-05-02	0	1	1758	1758	No	No	Si	Si	No	No	0	2457f4f34f53
10 - script PowerShell embebido	10	script PowerShell embebido	941	1	1	1	0	2026-05-02	2026-05-02	0	1	1124	1124	No	No	Si	Si	No	No	0	567a4aee3915
10 - script PowerShell embebido	10	script PowerShell embebido	942	1	1	1	0	2026-05-02	2026-05-02	0	1	1631	1631	No	No	Si	Si	No	No	0	c7d5846e7b08
10 - script PowerShell embebido	10	script PowerShell embebido	943	1	1	1	0	2026-05-02	2026-05-02	0	1	1707	1707	No	No	Si	Si	No	No	0	d03f97d9cca4
10 - script PowerShell embebido	10	script PowerShell embebido	944	1	1	1	0	2026-05-02	2026-05-02	1	0	1695	1695	No	No	Si	Si	No	No	0	5a78e15231f2
10 - script PowerShell embebido	10	script PowerShell embebido	945	1	1	1	0	2026-05-02	2026-05-02	0	1	1715	1715	No	No	Si	Si	No	No	0	3d38020b317f
10 - script PowerShell embebido	10	script PowerShell embebido	946	1	1	1	0	2026-05-02	2026-05-02	1	0	1715	1715	No	No	Si	Si	No	No	0	72475b30a977
10 - script PowerShell embebido	10	script PowerShell embebido	947	1	1	1	0	2026-05-02	2026-05-02	0	1	1710	1710	No	No	Si	Si	No	No	0	704774768830
10 - script PowerShell embebido	10	script PowerShell embebido	948	1	1	1	0	2026-05-02	2026-05-02	0	1	1722	1722	No	No	Si	Si	No	No	0	44c3dc258227
10 - script PowerShell embebido	10	script PowerShell embebido	949	1	1	1	0	2026-05-02	2026-05-02	0	1	2104	2104	No	No	No	Si	No	No	0	7b3bc4b2bfe7
10 - script PowerShell embebido	10	script PowerShell embebido	950	1	1	1	0	2026-05-02	2026-05-02	1	0	2118	2118	No	No	No	Si	No	No	0	aed4624a7b77
10 - script PowerShell embebido	10	script PowerShell embebido	951	1	1	1	0	2026-05-02	2026-05-02	1	0	2120	2120	No	No	No	Si	No	No	0	3c5045b8e658
10 - script PowerShell embebido	10	script PowerShell embebido	952	1	1	1	0	2026-05-02	2026-05-02	1	0	2120	2120	No	No	No	Si	No	No	0	ed38842f50b9
10 - script PowerShell embebido	10	script PowerShell embebido	953	1	1	1	0	2026-05-02	2026-05-02	1	0	2126	2126	No	No	No	Si	No	No	0	7f85257df598
10 - script PowerShell embebido	10	script PowerShell embebido	954	1	1	1	0	2026-05-02	2026-05-02	1	0	2120	2120	No	No	No	Si	No	No	0	e2eedad206fc
10 - script PowerShell embebido	10	script PowerShell embebido	955	1	1	1	0	2026-05-02	2026-05-02	1	0	2116	2116	No	No	No	Si	No	No	0	47cf0d53d9a8
10 - script PowerShell embebido	10	script PowerShell embebido	956	1	1	1	0	2026-05-02	2026-05-02	1	0	2112	2112	No	No	No	Si	No	No	0	381dedf1a9d6
10 - script PowerShell embebido	10	script PowerShell embebido	957	1	1	1	0	2026-05-02	2026-05-02	0	1	2178	2178	No	No	No	Si	No	No	0	209c67e2d988
10 - script PowerShell embebido	10	script PowerShell embebido	958	1	1	1	0	2026-05-02	2026-05-02	0	1	4113	4113	No	No	No	Si	No	No	0	29884f73c7c5
10 - script PowerShell embebido	10	script PowerShell embebido	959	1	1	1	0	2026-05-01	2026-05-01	0	1	2010	2010	No	No	No	Si	No	No	0	8974157b43a5
10 - script PowerShell embebido	10	script PowerShell embebido	960	1	1	1	0	2026-05-01	2026-05-01	0	1	2020	2020	No	No	No	Si	No	No	0	a8446a564ac3
10 - script PowerShell embebido	10	script PowerShell embebido	961	1	1	1	0	2026-05-01	2026-05-01	0	1	4134	4134	No	No	No	Si	No	No	0	124a95baece6
10 - script PowerShell embebido	10	script PowerShell embebido	962	1	1	1	0	2026-05-01	2026-05-01	0	1	2034	2034	No	No	No	Si	No	No	0	1fe89156f1a7
10 - script PowerShell embebido	10	script PowerShell embebido	963	1	1	1	0	2026-05-01	2026-05-01	1	0	2008	2008	No	No	No	Si	No	No	0	eff88db947ed
10 - script PowerShell embebido	10	script PowerShell embebido	964	1	1	1	0	2026-05-01	2026-05-01	0	1	2012	2012	No	No	No	Si	No	No	0	f90f42968def
10 - script PowerShell embebido	10	script PowerShell embebido	965	1	1	1	0	2026-05-01	2026-05-01	1	0	2034	2034	No	No	No	Si	No	No	0	994dec469f42
10 - script PowerShell embebido	10	script PowerShell embebido	966	1	1	1	0	2026-05-01	2026-05-01	0	1	2028	2028	No	No	No	Si	No	No	0	466dfad6eb27
10 - script PowerShell embebido	10	script PowerShell embebido	967	1	1	1	0	2026-05-01	2026-05-01	1	0	2042	2042	No	No	No	Si	No	No	0	c41ecaefc062
10 - script PowerShell embebido	10	script PowerShell embebido	968	1	1	1	0	2026-05-01	2026-05-01	0	1	46	46	No	Si	No	Si	No	No	0	361eba367fbe
10 - script PowerShell embebido	10	script PowerShell embebido	969	1	1	1	0	2026-05-01	2026-05-01	1	0	2016	2016	No	No	No	Si	No	No	0	a918f8094463
10 - script PowerShell embebido	10	script PowerShell embebido	970	1	1	1	0	2026-05-01	2026-05-01	0	1	2012	2012	No	No	No	Si	No	No	0	58f897b656aa
10 - script PowerShell embebido	10	script PowerShell embebido	971	1	1	1	0	2026-05-01	2026-05-01	0	1	2018	2018	No	No	No	Si	No	No	0	51415123f894
10 - script PowerShell embebido	10	script PowerShell embebido	972	1	1	1	0	2026-05-01	2026-05-01	1	0	2010	2010	No	No	No	Si	No	No	0	a02225c1aa3e
10 - script PowerShell embebido	10	script PowerShell embebido	973	1	1	1	0	2026-05-01	2026-05-01	1	0	2012	2012	No	No	No	Si	No	No	0	7bd702160398
10 - script PowerShell embebido	10	script PowerShell embebido	974	1	1	1	0	2026-05-01	2026-05-01	0	1	2008	2008	No	No	No	Si	No	No	0	c9dab8bc4eb2
10 - script PowerShell embebido	10	script PowerShell embebido	975	1	1	1	0	2026-04-15	2026-04-15	0	1	2022	2022	No	No	No	Si	No	No	0	e2d070cb11e6
10 - script PowerShell embebido	10	script PowerShell embebido	976	1	1	1	0	2026-04-15	2026-04-15	1	0	2030	2030	No	No	No	Si	No	No	0	dd615cce2022
10 - script PowerShell embebido	10	script PowerShell embebido	977	1	1	1	0	2026-04-15	2026-04-15	0	1	2028	2028	No	No	No	Si	No	No	0	241434de6ae8
10 - script PowerShell embebido	10	script PowerShell embebido	978	1	1	1	0	2026-04-15	2026-04-15	1	0	2022	2022	No	No	No	Si	No	No	0	6044b802efc8
10 - script PowerShell embebido	10	script PowerShell embebido	979	1	1	1	0	2026-04-15	2026-04-15	0	1	2026	2026	No	No	No	Si	No	No	0	d02bfdc639a4
10 - script PowerShell embebido	10	script PowerShell embebido	980	1	1	1	0	2026-04-15	2026-04-15	1	0	2030	2030	No	No	No	Si	No	No	0	5e28b29a2737
10 - script PowerShell embebido	10	script PowerShell embebido	981	1	1	1	0	2026-04-15	2026-04-15	0	1	2036	2036	No	No	No	Si	No	No	0	2676e4fba360
10 - script PowerShell embebido	10	script PowerShell embebido	982	1	1	1	0	2026-04-15	2026-04-15	0	1	2050	2050	No	No	No	Si	No	No	0	c91f4007b85b
10 - script PowerShell embebido	10	script PowerShell embebido	983	1	1	1	0	2026-04-15	2026-04-15	1	0	2030	2030	No	No	No	Si	No	No	0	c4b3cf2050be
10 - script PowerShell embebido	10	script PowerShell embebido	984	1	1	1	0	2026-04-15	2026-04-15	1	0	2030	2030	No	No	No	Si	No	No	0	9b5fa025b505
10 - script PowerShell embebido	10	script PowerShell embebido	985	1	1	1	0	2026-04-15	2026-04-15	0	1	2024	2024	No	No	No	Si	No	No	0	2e2eb12bd964
10 - script PowerShell embebido	10	script PowerShell embebido	986	1	1	1	0	2026-04-15	2026-04-15	0	1	2028	2028	No	No	No	Si	No	No	0	4efddf97fc3c
10 - script PowerShell embebido	10	script PowerShell embebido	987	1	1	1	0	2026-04-15	2026-04-15	0	1	2040	2040	No	No	No	Si	No	No	0	28ae6455ba26
10 - script PowerShell embebido	10	script PowerShell embebido	988	1	1	1	0	2026-04-15	2026-04-15	1	0	2024	2024	No	No	No	Si	No	No	0	8e6351384fbe
10 - script PowerShell embebido	10	script PowerShell embebido	989	1	1	1	0	2026-04-15	2026-04-15	1	0	2016	2016	No	No	No	Si	No	No	0	78cc79aa9d15
10 - script PowerShell embebido	10	script PowerShell embebido	990	1	1	1	0	2026-04-15	2026-04-15	1	0	2046	2046	No	No	No	Si	No	No	0	89bd1fadaf90
10 - script PowerShell embebido	10	script PowerShell embebido	991	1	1	1	0	2026-04-15	2026-04-15	1	0	2034	2034	No	No	No	Si	No	No	0	93e9838e7566
10 - script PowerShell embebido	10	script PowerShell embebido	992	1	1	1	0	2026-04-15	2026-04-15	1	0	2028	2028	No	No	No	Si	No	No	0	14783c8c8dc6
10 - script PowerShell embebido	10	script PowerShell embebido	993	1	1	1	0	2026-04-15	2026-04-15	0	1	4218	4218	No	No	No	Si	No	No	0	002d534cda69
10 - script PowerShell embebido	10	script PowerShell embebido	994	1	1	1	0	2026-04-15	2026-04-15	1	0	2030	2030	No	No	No	Si	No	No	0	6ec047c1eee6
10 - script PowerShell embebido	10	script PowerShell embebido	995	1	1	1	0	2026-04-15	2026-04-15	1	0	2038	2038	No	No	No	Si	No	No	0	fc1daa9f8c4f
10 - script PowerShell embebido	10	script PowerShell embebido	996	1	1	1	0	2026-04-15	2026-04-15	0	1	2060	2060	No	No	No	Si	No	No	0	f767a8e1a870
10 - script PowerShell embebido	10	script PowerShell embebido	997	1	1	1	0	2026-04-15	2026-04-15	1	0	2030	2030	No	No	No	Si	No	No	0	5b755827c6ae
10 - script PowerShell embebido	10	script PowerShell embebido	998	1	1	1	0	2026-04-15	2026-04-15	1	0	2020	2020	No	No	No	Si	No	No	0	ad4e222c6c4c
10 - script PowerShell embebido	10	script PowerShell embebido	999	1	1	1	0	2026-04-15	2026-04-15	1	0	2048	2048	No	No	No	Si	No	No	0	bf43ddd2253b
10 - script PowerShell embebido	10	script PowerShell embebido	1000	1	1	1	0	2026-04-15	2026-04-15	0	1	2036	2036	No	No	No	Si	No	No	0	f008edbb86e6
10 - script PowerShell embebido	10	script PowerShell embebido	1001	1	1	1	0	2026-04-15	2026-04-15	0	1	2032	2032	No	No	No	Si	No	No	0	61071f53d345
10 - script PowerShell embebido	10	script PowerShell embebido	1002	1	1	1	0	2026-04-15	2026-04-15	0	1	2016	2016	No	No	No	Si	No	No	0	9acda0b2839f
10 - script PowerShell embebido	10	script PowerShell embebido	1003	1	1	1	0	2026-04-15	2026-04-15	0	1	2040	2040	No	No	No	Si	No	No	0	b760efe0404f
10 - script PowerShell embebido	10	script PowerShell embebido	1004	1	1	1	0	2026-04-15	2026-04-15	1	0	2016	2016	No	No	No	Si	No	No	0	e82f0cfec9a8
10 - script PowerShell embebido	10	script PowerShell embebido	1005	1	1	1	0	2026-04-15	2026-04-15	1	0	2038	2038	No	No	No	Si	No	No	0	ab4add38fa47
10 - script PowerShell embebido	10	script PowerShell embebido	1006	1	1	1	0	2026-04-15	2026-04-15	1	0	2026	2026	No	No	No	Si	No	No	0	5b894728394e
10 - script PowerShell embebido	10	script PowerShell embebido	1007	1	1	1	0	2026-04-15	2026-04-15	1	0	2040	2040	No	No	No	Si	No	No	0	87332c1d7dc0
10 - script PowerShell embebido	10	script PowerShell embebido	1008	1	1	1	0	2026-04-15	2026-04-15	0	1	4248	4248	No	No	No	Si	No	No	0	eb4f83dc3818
10 - script PowerShell embebido	10	script PowerShell embebido	1009	1	1	1	0	2026-04-15	2026-04-15	1	0	2054	2054	No	No	No	Si	No	No	0	ecc728670631
10 - script PowerShell embebido	10	script PowerShell embebido	1010	1	1	1	0	2026-04-15	2026-04-15	0	1	2026	2026	No	No	No	Si	No	No	0	9a953494bdd4
10 - script PowerShell embebido	10	script PowerShell embebido	1011	1	1	1	0	2026-04-15	2026-04-15	1	0	2050	2050	No	No	No	Si	No	No	0	2207bfad79a2
10 - script PowerShell embebido	10	script PowerShell embebido	1012	1	1	1	0	2026-04-15	2026-04-15	0	1	2016	2016	No	No	No	Si	No	No	0	17566c341b72
10 - script PowerShell embebido	10	script PowerShell embebido	1013	1	1	1	0	2026-04-15	2026-04-15	1	0	2026	2026	No	No	No	Si	No	No	0	77d5b2e51064
10 - script PowerShell embebido	10	script PowerShell embebido	1014	1	1	1	0	2026-04-15	2026-04-15	1	0	2032	2032	No	No	No	Si	No	No	0	e6926f29b79f
10 - script PowerShell embebido	10	script PowerShell embebido	1015	1	1	1	0	2026-04-15	2026-04-15	0	1	2016	2016	No	No	No	Si	No	No	0	f9327a097f35
10 - script PowerShell embebido	10	script PowerShell embebido	1016	1	1	1	0	2026-04-15	2026-04-15	0	1	2036	2036	No	No	No	Si	No	No	0	d7704b1f4256
10 - script PowerShell embebido	10	script PowerShell embebido	1017	1	1	1	0	2026-04-14	2026-04-14	0	1	2032	2032	No	No	No	Si	No	No	0	3632826582d1
10 - script PowerShell embebido	10	script PowerShell embebido	1018	1	1	1	0	2026-04-14	2026-04-14	0	1	2030	2030	No	No	No	Si	No	No	0	fdda685c427f
10 - script PowerShell embebido	10	script PowerShell embebido	1019	1	1	1	0	2026-04-14	2026-04-14	0	1	2016	2016	No	No	No	Si	No	No	0	02c71ecc2928
10 - script PowerShell embebido	10	script PowerShell embebido	1020	1	1	1	0	2026-04-14	2026-04-14	1	0	2030	2030	No	No	No	Si	No	No	0	a2b203a0ce7a
10 - script PowerShell embebido	10	script PowerShell embebido	1021	1	1	1	0	2026-04-14	2026-04-14	0	1	2016	2016	No	No	No	Si	No	No	0	976576a60a84
10 - script PowerShell embebido	10	script PowerShell embebido	1022	1	1	1	0	2026-04-14	2026-04-14	1	0	2048	2048	No	No	No	Si	No	No	0	332fb78c526d
10 - script PowerShell embebido	10	script PowerShell embebido	1023	1	1	1	0	2026-04-14	2026-04-14	1	0	2034	2034	No	No	No	Si	No	No	0	dd987da02709
10 - script PowerShell embebido	10	script PowerShell embebido	1024	1	1	1	0	2026-04-14	2026-04-14	0	1	4274	4274	No	No	No	Si	No	No	0	1dca779e0e93
10 - script PowerShell embebido	10	script PowerShell embebido	1026	1	1	1	0	2026-04-14	2026-04-14	0	1	4105	4105	No	No	No	Si	No	No	0	f7d1126f6f90
10 - script PowerShell embebido	10	script PowerShell embebido	1027	1	1	1	0	2026-04-13	2026-04-13	0	1	4253	4253	No	No	No	Si	No	No	0	dac9d1acd89d
10 - script PowerShell embebido	10	script PowerShell embebido	1028	1	1	1	0	2026-04-13	2026-04-13	0	1	4128	4128	No	No	No	Si	No	No	0	5f52dda9ecfb
10 - script PowerShell embebido	10	script PowerShell embebido	1029	1	1	1	0	2026-04-13	2026-04-13	0	1	4141	4141	No	No	No	Si	No	No	0	b7acdcca43d2
10 - script PowerShell embebido	10	script PowerShell embebido	1030	1	1	1	0	2026-04-13	2026-04-13	0	1	4154	4154	No	No	No	Si	No	No	0	683e4a2cd9bc
10 - script PowerShell embebido	10	script PowerShell embebido	1031	1	1	1	0	2026-04-13	2026-04-13	0	1	4110	4110	No	No	No	Si	No	No	0	1e041d665842
10 - script PowerShell embebido	10	script PowerShell embebido	1032	1	1	1	0	2026-04-12	2026-04-12	0	1	4137	4137	No	No	No	Si	No	No	0	08ef7efc2f3b
10 - script PowerShell embebido	10	script PowerShell embebido	1033	1	1	1	0	2026-04-12	2026-04-12	0	1	4145	4145	No	No	No	Si	No	No	0	192cee3b3d7a
10 - script PowerShell embebido	10	script PowerShell embebido	1034	1	1	1	0	2026-04-12	2026-04-12	0	1	4208	4208	No	No	No	Si	No	No	0	ce44278f5ec8
10 - script PowerShell embebido	10	script PowerShell embebido	1035	1	1	1	0	2026-04-12	2026-04-12	0	1	4152	4152	No	No	No	Si	No	No	0	77c61f02d472
10 - script PowerShell embebido	10	script PowerShell embebido	1036	1	1	1	0	2026-04-12	2026-04-12	0	1	4123	4123	No	No	No	Si	No	No	0	e0e4ce7ae249
10 - script PowerShell embebido	10	script PowerShell embebido	1037	1	1	1	0	2026-04-12	2026-04-12	0	1	4142	4142	No	No	No	Si	No	No	0	2b02dc39a694
10 - script PowerShell embebido	10	script PowerShell embebido	1038	1	1	1	0	2026-04-12	2026-04-12	0	1	4217	4217	No	No	No	Si	No	No	0	0c51debf6f47
10 - script PowerShell embebido	10	script PowerShell embebido	1039	1	1	1	0	2026-04-11	2026-04-11	0	1	178	178	No	No	Si	Si	No	No	0	73ba266361ca
10 - script PowerShell embebido	10	script PowerShell embebido	1040	1	1	1	0	2026-04-11	2026-04-11	0	1	4239	4239	No	No	No	Si	No	No	0	c75997f4e54e
10 - script PowerShell embebido	10	script PowerShell embebido	1041	1	1	1	0	2026-04-11	2026-04-11	0	1	4109	4109	No	No	No	Si	No	No	0	4765e7a9766c
10 - script PowerShell embebido	10	script PowerShell embebido	1042	1	1	1	0	2026-04-11	2026-04-11	0	1	2016	2016	No	No	No	Si	No	No	0	963696f018e5
10 - script PowerShell embebido	10	script PowerShell embebido	1043	1	1	1	0	2026-04-07	2026-04-07	1	0	4137	4137	No	No	No	Si	No	No	0	348438d90564
10 - script PowerShell embebido	10	script PowerShell embebido	1044	1	1	1	0	2026-04-07	2026-04-07	0	1	2016	2016	No	No	No	Si	No	No	0	38325f3ee53d
10 - script PowerShell embebido	10	script PowerShell embebido	1047	1	1	1	0	2026-04-07	2026-04-07	1	0	2026	2026	No	No	No	Si	No	No	0	ccfef59e80d8
10 - script PowerShell embebido	10	script PowerShell embebido	1048	1	1	1	0	2026-04-07	2026-04-07	0	1	4166	4166	No	No	No	Si	No	No	0	89c058ea7069
10 - script PowerShell embebido	10	script PowerShell embebido	1049	1	1	1	0	2026-04-07	2026-04-07	0	1	2012	2012	No	No	No	Si	No	No	0	f3f98ede7c7b
10 - script PowerShell embebido	10	script PowerShell embebido	1050	1	1	1	0	2026-04-06	2026-04-06	1	0	38	38	No	No	No	No	No	No	0	202f74e4e132
10 - script PowerShell embebido	10	script PowerShell embebido	1053	1	1	1	0	2026-04-06	2026-04-06	1	0	2020	2020	No	No	No	Si	No	No	0	84cb0614e5de
10 - script PowerShell embebido	10	script PowerShell embebido	1054	1	1	1	0	2026-04-06	2026-04-06	0	1	4132	4132	No	No	No	Si	No	No	0	3380ea8a2388
10 - script PowerShell embebido	10	script PowerShell embebido	1058	1	1	1	0	2026-04-06	2026-04-06	0	1	4126	4126	No	No	No	Si	No	No	0	473f3ea9d34b
10 - script PowerShell embebido	10	script PowerShell embebido	1059	1	1	1	0	2026-04-06	2026-04-06	0	1	4132	4132	No	No	No	Si	No	No	0	e83793e8e7ba
10 - script PowerShell embebido	10	script PowerShell embebido	1060	1	1	1	0	2026-04-05	2026-04-05	0	1	2006	2006	No	No	No	Si	No	No	0	e21bf26adf68
10 - script PowerShell embebido	10	script PowerShell embebido	1061	1	1	1	0	2026-04-05	2026-04-05	1	0	2026	2026	No	No	No	Si	No	No	0	7e538f1efea0
10 - script PowerShell embebido	10	script PowerShell embebido	1062	1	1	1	0	2026-04-05	2026-04-05	0	1	2152	2152	No	No	No	Si	No	No	0	1ff207840136
10 - script PowerShell embebido	10	script PowerShell embebido	1063	1	1	1	0	2026-04-05	2026-04-05	1	0	4195	4195	No	No	No	Si	No	No	0	8ec007ccc923
10 - script PowerShell embebido	10	script PowerShell embebido	1064	1	1	1	0	2026-04-04	2026-04-04	0	1	4131	4131	No	No	No	Si	No	No	0	2d5922e8af7f
10 - script PowerShell embebido	10	script PowerShell embebido	1065	1	1	1	0	2026-04-04	2026-04-04	0	1	2016	2016	No	No	No	Si	No	No	0	51db40d9dbbc
10 - script PowerShell embebido	10	script PowerShell embebido	1066	1	1	1	0	2026-04-04	2026-04-04	0	1	4136	4136	No	No	No	Si	No	No	0	9b4a72464d3c
10 - script PowerShell embebido	10	script PowerShell embebido	1067	1	1	1	0	2026-04-04	2026-04-04	0	1	4128	4128	No	No	No	Si	No	No	0	d63977c463ae
10 - script PowerShell embebido	10	script PowerShell embebido	1068	1	1	1	0	2026-04-04	2026-04-04	0	1	4136	4136	No	No	No	Si	No	No	0	65d4ebe3dca9
10 - script PowerShell embebido	10	script PowerShell embebido	1069	1	1	1	0	2026-04-04	2026-04-04	0	1	4099	4099	No	No	No	Si	No	No	0	f3ddc8728468
10 - script PowerShell embebido	10	script PowerShell embebido	1071	1	1	1	0	2026-04-04	2026-04-04	0	1	4071	4071	No	No	No	Si	No	No	0	561a70a60c27
10 - script PowerShell embebido	10	script PowerShell embebido	1072	1	1	1	0	2026-04-04	2026-04-04	0	1	2026	2026	No	No	No	Si	No	No	0	55f019f826eb
10 - script PowerShell embebido	10	script PowerShell embebido	1073	1	1	1	0	2026-04-03	2026-04-03	0	1	2016	2016	No	No	No	Si	No	No	0	4a2611445997
10 - script PowerShell embebido	10	script PowerShell embebido	1074	1	1	1	0	2026-04-03	2026-04-03	0	1	2014	2014	No	No	No	Si	No	No	0	b45b27250cbb
10 - script PowerShell embebido	10	script PowerShell embebido	1075	1	1	1	0	2026-04-03	2026-04-03	0	1	4129	4129	No	No	No	Si	No	No	0	7ef33f3dea43
10 - script PowerShell embebido	10	script PowerShell embebido	1077	1	1	1	0	2026-04-02	2026-04-02	0	1	4116	4116	No	No	No	Si	No	No	0	e66f6398cbb9
10 - script PowerShell embebido	10	script PowerShell embebido	1079	1	1	1	0	2026-04-02	2026-04-02	0	1	4144	4144	No	No	No	Si	No	No	0	9389eb4756c8
10 - script PowerShell embebido	10	script PowerShell embebido	1083	1	1	1	0	2026-04-02	2026-04-02	0	1	2018	2018	No	No	No	Si	No	No	0	ac1e4d7fd2a6
10 - script PowerShell embebido	10	script PowerShell embebido	1084	1	1	1	0	2026-04-02	2026-04-02	1	0	2024	2024	No	No	No	Si	No	No	0	6efb3be9cd4e
10 - script PowerShell embebido	10	script PowerShell embebido	1085	1	1	1	0	2026-04-01	2026-04-01	0	1	2014	2014	No	No	No	Si	No	No	0	b365e1e66d80
10 - script PowerShell embebido	10	script PowerShell embebido	1088	1	1	1	0	2026-03-31	2026-03-31	0	1	2016	2016	No	No	No	Si	No	No	0	ed3c51933cff
10 - script PowerShell embebido	10	script PowerShell embebido	1090	1	1	1	0	2026-03-31	2026-03-31	0	1	4137	4137	No	No	No	Si	No	No	0	aafab1015d01
10 - script PowerShell embebido	10	script PowerShell embebido	1108	1	1	1	0	2026-03-31	2026-03-31	0	1	4118	4118	No	No	No	Si	No	No	0	d7ed331790f4
10 - script PowerShell embebido	10	script PowerShell embebido	1109	1	1	1	0	2026-03-30	2026-03-30	0	1	2016	2016	No	No	No	Si	No	No	0	21bb54d8c66c
10 - script PowerShell embebido	10	script PowerShell embebido	1110	1	1	1	0	2026-03-30	2026-03-30	0	1	2016	2016	No	No	No	Si	No	No	0	7f1650d451ae
10 - script PowerShell embebido	10	script PowerShell embebido	1111	1	1	1	0	2026-03-30	2026-03-30	0	1	2018	2018	No	No	No	Si	No	No	0	8b8531477139
10 - script PowerShell embebido	10	script PowerShell embebido	1112	1	1	1	0	2026-03-30	2026-03-30	0	1	4076	4076	No	No	No	Si	No	No	0	1e176a416d7b
10 - script PowerShell embebido	10	script PowerShell embebido	1113	1	1	1	0	2026-03-30	2026-03-30	0	1	2018	2018	No	No	No	Si	No	No	0	6465b3a4897a
10 - script PowerShell embebido	10	script PowerShell embebido	1115	1	1	1	0	2026-03-30	2026-03-30	0	1	2018	2018	No	No	No	Si	No	No	0	11261d0bbf6f
10 - script PowerShell embebido	10	script PowerShell embebido	1116	1	1	1	0	2026-03-29	2026-03-29	0	1	2012	2012	No	No	No	Si	No	No	0	9f69a31cc25e
10 - script PowerShell embebido	10	script PowerShell embebido	1122	1	1	1	0	2026-03-26	2026-03-26	0	1	4165	4165	No	No	No	Si	No	No	0	3406d1ca9674
10 - script PowerShell embebido	10	script PowerShell embebido	1123	1	1	1	0	2026-03-26	2026-03-26	1	0	4074	4074	No	No	No	Si	No	No	0	408963e284c4
10 - script PowerShell embebido	10	script PowerShell embebido	1125	1	1	1	0	2026-03-26	2026-03-26	0	1	4033	4033	No	No	No	Si	No	No	0	7dc5cc5945d7
10 - script PowerShell embebido	10	script PowerShell embebido	1128	1	1	1	0	2026-03-25	2026-03-25	0	1	4199	4199	No	No	No	Si	No	No	0	efe26517bca5
10 - script PowerShell embebido	10	script PowerShell embebido	1129	1	1	1	0	2026-03-25	2026-03-25	0	1	2020	2020	No	No	No	Si	No	No	0	68dbf5c6aea2
10 - script PowerShell embebido	10	script PowerShell embebido	1130	1	1	1	0	2026-03-25	2026-03-25	0	1	2020	2020	No	No	No	Si	No	No	0	f6ab9c2e16d5
10 - script PowerShell embebido	10	script PowerShell embebido	1131	1	1	1	0	2026-03-25	2026-03-25	0	1	2020	2020	No	No	No	Si	No	No	0	4b021997c9de
10 - script PowerShell embebido	10	script PowerShell embebido	1132	1	1	1	0	2026-03-25	2026-03-25	0	1	2020	2020	No	No	No	Si	No	No	0	c6adc4cb069c
10 - script PowerShell embebido	10	script PowerShell embebido	1133	1	1	1	0	2026-03-25	2026-03-25	0	1	2020	2020	No	No	No	Si	No	No	0	5d4dfbde3415
10 - script PowerShell embebido	10	script PowerShell embebido	1134	1	1	1	0	2026-03-25	2026-03-25	0	1	2103	2103	No	No	No	Si	No	No	0	e9b2cf0b4130
10 - script PowerShell embebido	10	script PowerShell embebido	1136	1	1	1	0	2026-03-24	2026-03-24	0	1	4121	4121	No	No	No	Si	No	No	0	e21e49d2ddc4
10 - script PowerShell embebido	10	script PowerShell embebido	1137	1	1	1	0	2026-03-24	2026-03-24	0	1	4071	4071	No	No	No	Si	No	No	0	dd9bc6270fa5
10 - script PowerShell embebido	10	script PowerShell embebido	1139	1	1	1	0	2026-03-24	2026-03-24	0	1	2010	2010	No	No	No	Si	No	No	0	b4fb7486b97b
10 - script PowerShell embebido	10	script PowerShell embebido	1140	1	1	1	0	2026-03-23	2026-03-23	0	1	4210	4210	No	No	No	Si	No	No	0	2550972c2ba6
10 - script PowerShell embebido	10	script PowerShell embebido	1141	1	1	1	0	2026-03-23	2026-03-23	0	1	4079	4079	No	No	No	Si	No	No	0	ee84a6d2f95b
10 - script PowerShell embebido	10	script PowerShell embebido	1142	1	1	1	0	2026-03-23	2026-03-23	0	1	4112	4112	No	No	No	Si	No	No	0	6bf00907927a
10 - script PowerShell embebido	10	script PowerShell embebido	1144	1	1	1	0	2026-03-22	2026-03-22	0	1	4232	4232	No	No	No	Si	No	No	0	f4c293d1e472
10 - script PowerShell embebido	10	script PowerShell embebido	1145	1	1	1	0	2026-03-21	2026-03-21	1	0	2051	2051	No	Si	No	Si	No	No	0	f6f4419b57f4
10 - script PowerShell embebido	10	script PowerShell embebido	1146	1	1	1	0	2026-03-21	2026-03-21	1	0	2020	2020	No	No	No	Si	No	No	0	596a5e8bfcf8
10 - script PowerShell embebido	10	script PowerShell embebido	1147	1	1	1	0	2026-03-21	2026-03-21	0	1	2020	2020	No	No	No	Si	No	No	0	cb61cd8db4c7
10 - script PowerShell embebido	10	script PowerShell embebido	1152	1	1	1	0	2026-03-16	2026-03-16	0	1	2024	2024	No	No	No	Si	No	No	0	489ae0de96d2
10 - script PowerShell embebido	10	script PowerShell embebido	1153	1	1	1	0	2026-03-16	2026-03-16	1	0	2024	2024	No	No	No	Si	No	No	0	5893cfaab742
10 - script PowerShell embebido	10	script PowerShell embebido	1161	1	1	1	0	2026-03-15	2026-03-15	0	1	48	48	No	Si	No	Si	No	No	0	377402608721
10 - script PowerShell embebido	10	script PowerShell embebido	1180	1	1	1	0	2026-03-14	2026-03-14	0	1	2016	2016	No	No	No	Si	No	No	0	80f6cc970159
10 - script PowerShell embebido	10	script PowerShell embebido	1185	1	1	1	0	2026-03-12	2026-03-12	0	1	2024	2024	No	No	No	Si	No	No	0	95f0e405007c
10 - script PowerShell embebido	10	script PowerShell embebido	1187	1	1	1	0	2026-03-12	2026-03-12	0	1	2024	2024	No	No	No	Si	No	No	0	d270ece3c2ed
10 - script PowerShell embebido	10	script PowerShell embebido	1189	1	1	1	0	2026-03-12	2026-03-12	0	1	2024	2024	No	No	No	Si	No	No	0	387b1c31a58a
10 - script PowerShell embebido	10	script PowerShell embebido	1192	1	1	1	0	2026-03-12	2026-03-12	0	1	4090	4090	No	No	No	Si	No	No	0	7c98f2f39f21
10 - script PowerShell embebido	10	script PowerShell embebido	1193	1	1	1	0	2026-03-10	2026-03-10	0	1	1981	1981	No	No	No	Si	No	No	0	36de1c428193
10 - script PowerShell embebido	10	script PowerShell embebido	1198	1	1	1	0	2026-03-09	2026-03-09	1	0	2066	2066	No	No	No	Si	No	No	0	c71016d232d2
10 - script PowerShell embebido	10	script PowerShell embebido	1199	1	1	1	0	2026-03-09	2026-03-09	1	0	2060	2060	No	No	No	Si	No	No	0	a35f85ce7761
10 - script PowerShell embebido	10	script PowerShell embebido	1200	1	1	1	0	2026-03-09	2026-03-09	1	0	2058	2058	No	No	No	Si	No	No	0	85372e650e0d
10 - script PowerShell embebido	10	script PowerShell embebido	1201	1	1	1	0	2026-03-09	2026-03-09	1	0	2080	2080	No	No	No	Si	No	No	0	92209c6106e3
10 - script PowerShell embebido	10	script PowerShell embebido	1202	1	1	1	0	2026-03-09	2026-03-09	1	0	2084	2084	No	No	No	Si	No	No	0	0aa596181f19
10 - script PowerShell embebido	10	script PowerShell embebido	1203	1	1	1	0	2026-03-09	2026-03-09	1	0	2106	2106	No	No	No	Si	No	No	0	1e8b488db6ca
10 - script PowerShell embebido	10	script PowerShell embebido	1204	1	1	1	0	2026-03-09	2026-03-09	1	0	2081	2081	No	No	No	Si	No	No	0	d09d68e9bb7c
10 - script PowerShell embebido	10	script PowerShell embebido	1205	1	1	1	0	2026-03-09	2026-03-09	1	0	2090	2090	No	No	No	Si	No	No	0	1ca488c17264
10 - script PowerShell embebido	10	script PowerShell embebido	1206	1	1	1	0	2026-03-09	2026-03-09	1	0	2094	2094	No	No	No	Si	No	No	0	2f4775548a01
10 - script PowerShell embebido	10	script PowerShell embebido	1207	1	1	1	0	2026-03-09	2026-03-09	1	0	2084	2084	No	No	No	Si	No	No	0	0608b77fb1e7
10 - script PowerShell embebido	10	script PowerShell embebido	1208	1	1	1	0	2026-03-09	2026-03-09	1	0	2074	2074	No	No	No	Si	No	No	0	fa6641f36f37
10 - script PowerShell embebido	10	script PowerShell embebido	1209	1	1	1	0	2026-03-09	2026-03-09	1	0	2084	2084	No	No	No	Si	No	No	0	3bd825cf13b2
10 - script PowerShell embebido	10	script PowerShell embebido	1210	1	1	1	0	2026-03-09	2026-03-09	1	0	2085	2085	No	No	No	Si	No	No	0	ede28c6fbb93
10 - script PowerShell embebido	10	script PowerShell embebido	1211	1	1	1	0	2026-03-09	2026-03-09	1	0	2098	2098	No	No	No	Si	No	No	0	83baf2074aee
10 - script PowerShell embebido	10	script PowerShell embebido	1212	1	1	1	0	2026-03-09	2026-03-09	1	0	2086	2086	No	No	No	Si	No	No	0	9b33f0910338
10 - script PowerShell embebido	10	script PowerShell embebido	1213	1	1	1	0	2026-03-09	2026-03-09	1	0	2053	2053	No	No	No	Si	No	No	0	fcfce5832751
10 - script PowerShell embebido	10	script PowerShell embebido	1214	1	1	1	0	2026-03-09	2026-03-09	1	0	2100	2100	No	No	No	Si	No	No	0	133d913e7cec
10 - script PowerShell embebido	10	script PowerShell embebido	1215	1	1	1	0	2026-03-09	2026-03-09	1	0	2077	2077	No	No	No	Si	No	No	0	4a503f361de9
10 - script PowerShell embebido	10	script PowerShell embebido	1216	1	1	1	0	2026-03-09	2026-03-09	1	0	2084	2084	No	No	No	Si	No	No	0	40a24982357f
10 - script PowerShell embebido	10	script PowerShell embebido	1217	1	1	1	0	2026-03-09	2026-03-09	1	0	3772	3772	No	No	No	Si	No	No	0	465496635ef4
10 - script PowerShell embebido	10	script PowerShell embebido	1218	1	1	1	0	2026-03-09	2026-03-09	1	0	2080	2080	No	No	No	Si	No	No	0	8c82ed23d57d
10 - script PowerShell embebido	10	script PowerShell embebido	1219	1	1	1	0	2026-03-09	2026-03-09	0	1	2102	2102	No	No	No	Si	No	No	0	6426b53edb92
10 - script PowerShell embebido	10	script PowerShell embebido	1220	1	1	1	0	2026-03-09	2026-03-09	1	0	2105	2105	No	No	No	Si	No	No	0	985e4dfad366
10 - script PowerShell embebido	10	script PowerShell embebido	1221	1	1	1	0	2026-03-09	2026-03-09	1	0	2103	2103	No	No	No	Si	No	No	0	ec20e513cd6b
10 - script PowerShell embebido	10	script PowerShell embebido	1222	1	1	1	0	2026-03-09	2026-03-09	1	0	2061	2061	No	No	No	Si	No	No	0	82f0accc2ae2
10 - script PowerShell embebido	10	script PowerShell embebido	1230	1	1	1	0	2026-03-06	2026-03-06	0	1	41	41	No	Si	No	Si	No	No	0	f9bd4d2799df
10 - script PowerShell embebido	10	script PowerShell embebido	1300	1	1	1	0	2026-01-04	2026-01-04	1	0	46	46	No	No	No	No	No	No	0	6f29e9f66b80
10 - script PowerShell embebido	10	script PowerShell embebido	1372	1	1	1	0	2025-10-15	2025-10-15	1	0	854	854	No	Si	Si	Si	No	No	0	0d4a0e67fac4
10 - script PowerShell embebido	10	script PowerShell embebido	1379	1	1	1	0	2025-10-09	2025-10-09	0	1	191	191	No	Si	Si	Si	No	No	0	4342426ffa42
10 - script PowerShell embebido	10	script PowerShell embebido	1380	1	1	1	0	2025-10-09	2025-10-09	0	1	183	183	No	Si	Si	Si	No	No	0	1fefac245bf5
10 - script PowerShell embebido	10	script PowerShell embebido	1409	1	1	1	0	2025-09-07	2025-09-07	0	1	27	27	No	Si	No	Si	No	No	0	d59db1861bf5
10 - script PowerShell embebido	10	script PowerShell embebido	1461	1	1	1	0	2025-08-24	2025-08-24	1	0	82	82	No	Si	No	Si	No	No	0	53388642f669
10 - script PowerShell embebido	10	script PowerShell embebido	1462	1	1	1	0	2025-08-24	2025-08-24	1	0	81	81	No	Si	No	Si	No	No	0	bc6596785952
11 - shell Unix (bash/ssh/curl)	11	shell Unix (bash/ssh/curl)	47	8	8	8	0	2025-10-10	2025-11-02	8	0	89	89	No	Si	No	No	No	No	0	8322770ef8d7
11 - shell Unix (bash/ssh/curl)	11	shell Unix (bash/ssh/curl)	68	5	5	5	0	2025-08-23	2025-08-26	4	1	97	97	No	Si	No	No	No	No	0	d293dafefe20
11 - shell Unix (bash/ssh/curl)	11	shell Unix (bash/ssh/curl)	160	2	2	2	0	2025-08-21	2025-08-23	2	0	93	93	No	Si	No	No	No	No	0	7ec8ea67af81
11 - shell Unix (bash/ssh/curl)	11	shell Unix (bash/ssh/curl)	340	1	1	1	0	2026-06-30	2026-06-30	0	1	107	107	No	No	No	No	No	No	0	e3409e143ef9
11 - shell Unix (bash/ssh/curl)	11	shell Unix (bash/ssh/curl)	611	1	1	1	0	2026-05-27	2026-05-27	1	0	176	176	No	Si	Si	No	No	No	0	193182b035c3
11 - shell Unix (bash/ssh/curl)	11	shell Unix (bash/ssh/curl)	1349	1	1	1	0	2025-11-13	2025-11-13	1	0	97	97	No	Si	No	No	No	No	0	d0fbb5d4881a
12 - solo URL / dominio	12	solo URL / dominio	74	4	4	4	0	2026-04-12	2026-04-13	0	4	73	73	No	Si	No	No	No	No	0	7b9f8a9e3250
12 - solo URL / dominio	12	solo URL / dominio	162	1	1	1	0	2026-07-29	2026-07-29	1	0	28	28	No	Si	No	No	No	No	0	01880d08684a
12 - solo URL / dominio	12	solo URL / dominio	167	1	1	1	0	2026-07-29	2026-07-29	1	0	20	20	No	No	No	No	No	No	0	574ef1932888
12 - solo URL / dominio	12	solo URL / dominio	173	1	1	1	0	2026-07-27	2026-07-27	1	0	23	23	No	Si	No	No	No	No	0	e200055bfb0c
12 - solo URL / dominio	12	solo URL / dominio	246	1	1	1	0	2026-07-21	2026-07-21	0	1	28	28	No	Si	No	No	No	No	0	c70ae6287559
12 - solo URL / dominio	12	solo URL / dominio	363	1	1	1	0	2026-06-26	2026-06-26	1	0	28	28	No	Si	No	No	No	No	0	e14447bf62ca
12 - solo URL / dominio	12	solo URL / dominio	552	1	1	1	0	2026-06-08	2026-06-08	1	0	28	28	No	Si	No	No	No	No	0	f8421c6057aa
12 - solo URL / dominio	12	solo URL / dominio	596	1	1	1	0	2026-06-08	2026-06-08	0	1	65	65	No	Si	No	No	No	No	0	2bedc1ae7af8
12 - solo URL / dominio	12	solo URL / dominio	606	1	1	1	0	2026-05-30	2026-05-30	1	0	22	22	No	Si	No	No	No	No	0	6518267be54d
12 - solo URL / dominio	12	solo URL / dominio	1351	1	1	1	0	2025-11-11	2025-11-11	0	1	124	124	No	Si	No	No	No	No	0	68bd7b6d408e
13 - echo	13	echo	1169	1	1	1	0	2026-03-15	2026-03-15	0	1	106	106	No	No	Si	No	No	No	0	8f9d45c9fd38
13 - echo	13	echo	1181	1	1	1	0	2026-03-14	2026-03-14	0	1	114	114	No	No	Si	No	No	No	0	251977dc332b
13 - echo	13	echo	1183	1	1	1	0	2026-03-13	2026-03-13	0	1	108	108	No	No	Si	No	No	No	0	2dbf1aadb8ce
13 - echo	13	echo	1245	1	1	1	0	2026-02-21	2026-02-21	1	0	317	317	No	No	No	No	No	No	0	ead85f73d72b
13 - echo	13	echo	1296	1	1	1	0	2026-01-04	2026-01-04	1	0	106	106	No	No	Si	No	No	No	0	5d1e33f0e18c
99 - otro	99	otro	41	9	9	9	0	2026-05-07	2026-07-22	9	0	10	10	No	No	No	No	No	No	0	6323dd1f0769
99 - otro	99	otro	137	2	2	1	1	2025-08-21	2026-01-04	0	2	15	15	No	No	No	No	No	No	0	ae7f30401a46
99 - otro	99	otro	236	1	1	1	0	2026-07-24	2026-07-24	0	1	7	7	No	No	No	No	No	No	0	2fd68fd91e61
99 - otro	99	otro	868	1	1	1	0	2026-05-06	2026-05-06	1	0	217	217	No	No	No	Si	Si	No	0	6d889571d9f4
99 - otro	99	otro	869	1	1	1	0	2026-05-06	2026-05-06	1	0	215	215	No	No	No	Si	Si	No	0	3a9f499715b6
99 - otro	99	otro	1051	1	1	0	0			0	0	3	3	No	No	No	No	No	No	1	eb77695c1d1c
99 - otro	99	otro	1052	1	1	0	0			0	0	1	1	No	No	No	No	No	No	1	440cc4c91ea6
99 - otro	99	otro	1195	1	1	1	0	2026-03-10	2026-03-10	0	1	3	3	No	No	No	No	No	No	0	716716ea53da
99 - otro	99	otro	1197	1	1	1	0	2026-03-09	2026-03-09	0	1	61	61	No	Si	No	No	No	No	0	08f76028fc85
99 - otro	99	otro	1293	1	1	1	0	2026-01-09	2026-01-09	0	1	9	9	No	No	No	No	No	No	0	2ee1838e9c5d
99 - otro	99	otro	1310	1	1	1	0	2025-12-13	2025-12-13	0	1	18	18	No	No	No	No	No	No	0	d3f0e7a219bc
99 - otro	99	otro	1314	1	1	1	0	2025-12-09	2025-12-09	0	1	6	6	No	No	No	No	No	No	0	7e1720ebb461
99 - otro	99	otro	1334	1	1	1	0	2025-11-26	2025-11-26	0	1	4	4	No	No	No	No	No	No	0	7b75090d62a6
99 - otro	99	otro	1335	1	1	1	0	2025-11-25	2025-11-25	0	1	94	94	No	Si	No	No	No	No	0	b89db4027f4d
99 - otro	99	otro	1393	1	1	1	0	2025-10-01	2025-10-01	1	0	184	184	No	Si	No	Si	Si	No	0	dc6df023a842
99 - otro	99	otro	1421	1	1	1	0	2025-09-02	2025-09-02	1	0	50	50	No	Si	No	No	No	No	0	6ef2198a3f5e
__GROUP_TSV__
)

META_TSV=$(cat <<'__META_TSV__'
BLOQUE	Clave	Valor
Procedencia	Fichero fuente	excel_clasificado.xlsx
Procedencia	Hoja fuente	clickfix_all (2)
Procedencia	Hoja descartada	clickfix_all (545 columnas, CSV mal importado)
Procedencia	Fichero secundario	excel.xlsx (sin registros unicos adicionales)
Procedencia	Generado	excel_limpio_estructural.xlsx (los originales no se modifican)
Dataset	Filas de datos en la fuente	4411
Dataset	Filas vacias (ZZ - fila vacia)	72
Dataset	Filas clasificadas en grupos	4339
Dataset	Categorias distintas (sin ZZ)	15
Dataset	Categorias distintas (con ZZ)	16
Dataset	Grupos distintos	1483
Dataset	Grupos unitarios (Repeticiones=1)	1322
Dataset	Dominios distintos (solo recuento)	3886
Dataset	Rango de fechas	2025-08-17 .. 2026-07-30
Dataset	Filas con senal de mal parseo	266
Dataset	Registros cuyo host es una IP literal	187
Integridad	Suma Num_Registros de Estructura	4339
Integridad	Suma Num_Registros de Resumen	4411
Integridad	Filas de Estructura (1 por grupo)	1483
Integridad	Repeticiones == Num_Registros en todos los grupos	Si
Integridad	Cada grupo pertenece a una sola categoria	Si
Integridad	Grupo y contenido de portapapeles	relacion 1 a 1 verificada en la fuente
Integridad	Filas unicas perdidas frente a excel.xlsx	0
Descartado	clickfix_all (ambos ficheros)	CSV volcado en 1 columna y desbordado a 542/545 columnas; 99-otro absorbe 3157 filas
Descartado	Clipboard Content	muestra cruda ejecutable; solo se derivan longitud y flags de tecnica
Descartado	Domain / URL / URLScan Link	infraestructura y enlaces a muestras; solo se conserva el recuento de dominios distintos
Descartado	Direccion del host	no se conserva ningun dominio ni ninguna IP, solo recuentos
Descartado	IP	10 valores, todos desbordamiento de columna (contenian el enlace de urlscan)
Descartado	Country / ASN / Server / Technologies / Tags / Page Title	1-2 valores no vacios en 4411 filas y todos residuo numerico de mal parseo; sin valor estructural
Descartado	Timestamp (hora exacta)	se conserva solo la fecha agregada minima y maxima por grupo
Constante	Malicious	Yes=4076; (vacio)=260; otros=3 (residuo de mal parseo, no se conserva el valor)
Constante	Clipboard Hijacked	Yes=4076; (vacio)=261; otros=2 (residuo de mal parseo, no se conserva el valor)
Constante	Has Hudson Rock	0=3567; (vacio)=769; otros=3 (residuo de mal parseo, no se conserva el valor)
Constante	Motivo de descarte	sin poder discriminante: se registran aqui y no por grupo
Politica	Contenido ejecutable	ninguno: ninguna orden de consola, guion, fragmento de codigo ni carga util
Politica	Red	ninguna URL, IP, dominio ni infraestructura externa concreta
Politica	Etiquetas de categoria	se conservan literales de la fuente por ser taxonomia; nombran tecnicas, no son ordenes ejecutables
Politica	Grupo_Hash	sha256 truncado a 12 hex del contenido con sal fija del dataset; clave de union local, no reconstruye ni permite buscar el contenido fuera
Politica	Flags de tecnica	heuristicos: se obtienen por coincidencia de patrones sobre el contenido, que despues se descarta; indican presencia, no prueban la tecnica
Politica	Verificacion	un validador independiente revisa cada celda de la salida y la coteja con la fuente
		
Hoja	Columna	Significado
Estructura	Categoria	Etiqueta de categoria tal cual aparece en la fuente. Clave de la taxonomia.
Estructura	Categoria_ID	Codigo numerico de la categoria (00-13, 99, ZZ). Orden estable para un menu.
Estructura	Categoria_Nombre	Parte descriptiva de la etiqueta, sin el codigo.
Estructura	Grupo	Identificador del grupo. Global y unico; cada grupo vive en una sola categoria.
Estructura	Repeticiones	Tamano del grupo segun la fuente (venia repetido en cada fila).
Estructura	Num_Registros	Tamano del grupo recontado aqui. Debe coincidir con Repeticiones.
Estructura	Dominios_Distintos	Numero de dominios distintos del grupo. Solo el recuento.
Estructura	Registros_Host_IP	Registros del grupo cuyo host es una IP literal en vez de un dominio. Solo el recuento; la IP no se conserva.
Estructura	Fecha_Primera	Fecha mas antigua observada en el grupo (sin hora).
Estructura	Fecha_Ultima	Fecha mas reciente observada en el grupo (sin hora).
Estructura	Captcha_Si	Registros del grupo con captcha detectado.
Estructura	Captcha_No	Registros del grupo sin captcha detectado.
Estructura	Longitud_Media	Longitud media en caracteres del contenido del grupo. Metrica, no contenido.
Estructura	Longitud_Max	Longitud maxima en caracteres del contenido del grupo.
Estructura	Multilinea	El contenido del grupo ocupa varias lineas.
Estructura	Descarga_Remota	Se observa una etapa de descarga o acceso remoto.
Estructura	Codificacion_Base64	Se observa codificacion en base64.
Estructura	Ofuscacion	Se observan tecnicas de ofuscacion de cadenas.
Estructura	Ejecucion_Oculta	Se observan indicadores de ejecucion sin ventana visible.
Estructura	Persistencia	Se observan indicadores de persistencia.
Estructura	Filas_Mal_Parseadas	Registros del grupo con senal de desbordamiento de campos en la fuente. Un host con IP literal no cuenta como error.
Estructura	Grupo_Hash	sha256 truncado y con sal del contenido del grupo. Clave de union estable; vacio si el grupo no tiene contenido.
Resumen	Num_Grupos	Grupos distintos dentro de la categoria.
Resumen	Num_Registros	Registros de la categoria.
Resumen	Pct_Registros	Porcentaje sobre las filas clasificadas en grupos.
Resumen	Dominios_Distintos	Dominios distintos de la categoria. Solo el recuento.
Resumen	Registros_Host_IP	Registros de la categoria cuyo host es una IP literal. Solo el recuento.
Resumen	Grupo_Mayor	Identificador del grupo mas numeroso de la categoria.
Resumen	Tam_Grupo_Mayor	Tamano de ese grupo.
Resumen	Grupos_Unitarios	Grupos de un solo registro.
Resumen	Longitud_Media	Longitud media ponderada por registros.
Resumen	Grupos_<flag>	Cuantos grupos de la categoria presentan ese flag.
Metadatos	BLOQUE, Clave, Valor	Procedencia, dataset, integridad, descartes, constantes y politica de saneado.
__META_TSV__
)

clear_screen() {
  if command -v clear >/dev/null 2>&1; then clear; fi
}

pause_screen() {
  printf "\nPulsa Enter para continuar..."
  IFS= read -r _
}

# ${var,,} y mapfile son de bash 4+; el bash de sistema de macOS (/bin/bash)
# es 3.2. lc() es el equivalente portable a minúsculas.
lc() {
  printf '%s' "$1" | tr '[:upper:]' '[:lower:]'
}

yes_no() {
  local prompt="${1:-Continuar?}"
  local ans
  printf "%s [s/N]: " "$prompt"
  IFS= read -r ans
  case "$(lc "$ans")" in s|si|sí|y|yes) return 0;; *) return 1;; esac
}

print_banner() {
  printf '%s\n' "============================================================"
  printf '%s\n' "  LAUNCHER DE LABORATORIO LOCAL"
  printf '%s\n' "============================================================"
  printf '%s\n' "Datos integrados: 16 categorías / 1483 grupos"
  printf '%s\n' "Ejecución: únicamente validaciones benignas/locales"
  printf '%s\n\n' "============================================================"
}

category_line_by_id() {
  local id="$1"
  printf '%s\n' "$SUMMARY_TSV" | awk -F '\t' -v id="$id" 'NR>1 && $2==id {print; exit}'
}

list_categories() {
  printf '%s\n' "$SUMMARY_TSV" | awk -F '\t' 'NR>1 {
    printf "%-3s %-46s  grupos:%-4s registros:%-5s\n", $2, $3, $4, $5
  }'
}

# read colapsa tabuladores consecutivos (son IFS whitespace) y pierde los
# campos vacíos del TSV, desalineando las columnas siguientes. Se convierte
# a un separador no-whitespace (0x1F) antes de partir la línea.
tsv_split() {
  local line="$1"; shift
  local IFS=$'\x1f'
  # shellcheck disable=SC2034
  read -r "$@" <<<"${line//$'\t'/$'\x1f'}"
}

show_category_detail() {
  local id="$1" line
  line="$(category_line_by_id "$id")"
  if [[ -z "$line" ]]; then
    printf "Categoría no encontrada.\n"
    return 1
  fi
  tsv_split "$line" cat cid name ngr nreg pct dom ip gmay tmay unit cap first last avg dl b64 obf hid pers bad
  printf "\nCategoría: %s - %s\n" "$cid" "$name"
  printf "Grupos: %s | Registros: %s | %s%% del dataset\n" "$ngr" "$nreg" "$pct"
  printf "Rango temporal: %s -> %s\n" "${first:-n/d}" "${last:-n/d}"
  printf "Grupos con flags: descarga=%s base64=%s ofuscación=%s oculta=%s persistencia=%s\n" \
    "$dl" "$b64" "$obf" "$hid" "$pers"
  printf "Filas mal parseadas registradas: %s\n" "$bad"
}

groups_for_category() {
  local id="$1"
  printf '%s\n' "$GROUP_TSV" | awk -F '\t' -v id="$id" 'NR>1 && $2==id'
}

show_group_detail_line() {
  local line="$1"
  tsv_split "$line" cat cid name grp rep nreg dom ip first last capy capn lavg lmax multi dl b64 obf hid pers bad hash
  printf "\n-------------------- GRUPO %s --------------------\n" "$grp"
  printf "Categoría: %s\n" "$cat"
  printf "Grupo/hash: %s / %s\n" "$grp" "${hash:-n/d}"
  printf "Repeticiones: %s | Registros: %s | Dominios distintos (recuento): %s | Host-IP (recuento): %s\n" \
    "$rep" "$nreg" "$dom" "$ip"
  printf "Fechas: %s -> %s\n" "${first:-n/d}" "${last:-n/d}"
  printf "Captcha: sí=%s no=%s | Longitud media/máx: %s/%s\n" "$capy" "$capn" "$lavg" "$lmax"
  printf "Flags: multilinea=%s descarga_remota=%s base64=%s ofuscacion=%s ejecucion_oculta=%s persistencia=%s\n" \
    "$multi" "$dl" "$b64" "$obf" "$hid" "$pers"
  printf "Filas mal parseadas: %s\n" "$bad"
}

mktemp_dir() {
  local d
  d="$(mktemp -d 2>/dev/null || mktemp -d -t lablauncher)"
  printf '%s' "$d"
}

run_base64_roundtrip() {
  command -v base64 >/dev/null 2>&1 || { printf "SKIP base64: utilidad no disponible\n"; return 0; }
  local src="LAB_OK_${RANDOM}" enc dec
  enc="$(printf '%s' "$src" | base64 | tr -d '\n')" || return 1
  dec="$(printf '%s' "$enc" | base64 -d 2>/dev/null)" || return 1
  [[ "$dec" == "$src" ]] || return 1
  printf "OK  codificación/decodificación Base64 local\n"
}

run_obfuscation_roundtrip() {
  # Transformación reversible de una cadena inocua; no ejecuta contenido decodificado.
  local src="LAB_OK" rev back
  rev="$(printf '%s' "$src" | rev 2>/dev/null || true)"
  if [[ -z "$rev" ]]; then
    printf "SKIP transformación: rev no disponible\n"
    return 0
  fi
  back="$(printf '%s' "$rev" | rev)"
  [[ "$back" == "$src" ]] || return 1
  printf "OK  transformación reversible local\n"
}

run_multiline_local() {
  local td f out
  td="$(mktemp_dir)" || return 1
  f="$td/test.sh"
  cat >"$f" <<'EOF'
#!/usr/bin/env sh
x="LAB"
y="_OK"
printf '%s%s\n' "$x" "$y"
EOF
  chmod +x "$f"
  out="$("$f" 2>&1)"
  rm -rf "$td"
  [[ "$out" == "LAB_OK" ]] || { printf "FAIL multilínea: %s\n" "$out"; return 1; }
  printf "OK  script multilínea local\n"
}

run_quiet_local() {
  local td marker pid
  td="$(mktemp_dir)" || return 1
  marker="$td/marker"
  sh -c 'printf LAB_OK >"$1"' sh "$marker" >/dev/null 2>&1 &
  pid=$!
  wait "$pid" || { rm -rf "$td"; return 1; }
  [[ "$(cat "$marker" 2>/dev/null)" == "LAB_OK" ]] || { rm -rf "$td"; return 1; }
  rm -rf "$td"
  printf "OK  subprocesso silencioso local\n"
}

run_local_http_fetch() {
  local td port pid i got=""
  command -v python3 >/dev/null 2>&1 || {
    printf "SKIP HTTP local: python3 no está disponible para levantar el servicio temporal\n"
    return 0
  }
  if command -v curl >/dev/null 2>&1; then
    :
  elif command -v wget >/dev/null 2>&1; then
    :
  else
    printf "SKIP HTTP local: no hay curl ni wget\n"
    return 0
  fi
  td="$(mktemp_dir)" || return 1
  printf 'LAB_OK\n' >"$td/marker.txt"
  port=$((18080 + RANDOM % 4000))
  python3 -m http.server "$port" --bind 127.0.0.1 --directory "$td" >"$td/http.log" 2>&1 &
  pid=$!
  for i in 1 2 3 4 5 6 7 8 9 10; do
    if kill -0 "$pid" 2>/dev/null; then
      if command -v curl >/dev/null 2>&1; then
        got="$(curl -fsS --max-time 2 "http://127.0.0.1:$port/marker.txt" 2>/dev/null || true)"
      else
        got="$(wget -qO- -T 2 "http://127.0.0.1:$port/marker.txt" 2>/dev/null || true)"
      fi
      [[ "$got" == "LAB_OK" ]] && break
    fi
    sleep 0.15
  done
  kill "$pid" 2>/dev/null || true
  wait "$pid" 2>/dev/null || true
  rm -rf "$td"
  [[ "$got" == "LAB_OK" ]] || { printf "FAIL HTTP local\n"; return 1; }
  printf "OK  descarga HTTP real desde 127.0.0.1\n"
}

run_category_linux() {
  local cid="$1"
  case "$cid" in
    08|10)
      if command -v pwsh >/dev/null 2>&1; then
        local out
        out="$(pwsh -NoProfile -NonInteractive -Command "Write-Output 'LAB_OK'" 2>&1)"
        [[ "$out" == *LAB_OK* ]] || return 1
        printf "OK  PowerShell local mediante pwsh\n"
      else
        printf "SKIP ejecutor PowerShell: pwsh no está instalado en Linux\n"
      fi
      ;;
    11)
      local out
      out="$(bash -c 'printf "LAB_OK\n"' 2>&1)" || return 1
      [[ "$out" == "LAB_OK" ]] || return 1
      printf "OK  bash local\n"
      ;;
    12)
      run_local_http_fetch || return 1
      ;;
    13)
      local out
      out="$(printf 'LAB_OK\n')"
      [[ "$out" == "LAB_OK" ]] || return 1
      printf "OK  echo/printf local\n"
      ;;
    01|02|03|04|05|06|07|09)
      printf "INFO esta categoría usa componentes nativos de Windows; ejecútala con launcher-windows.ps1\n"
      ;;
    00|99|ZZ)
      printf "INFO la taxonomía no define un ejecutor específico para esta categoría; se conserva íntegra como metadato.\n"
      ;;
    *)
      printf "INFO categoría sin ejecutor específico.\n"
      ;;
  esac
}

run_group_safe() {
  local line="$1"
  tsv_split "$line" cat cid name grp rep nreg dom ip first last capy capn lavg lmax multi dl b64 obf hid pers bad hash
  printf "\nResumen de ejecución:\n"
  printf "  Categoría: %s\n  Grupo: %s\n  Hash: %s\n  Modo: laboratorio local seguro\n" "$cat" "$grp" "${hash:-n/d}"
  yes_no "¿Ejecutar ahora?" || { printf "Cancelado.\n"; return 0; }
  printf "\n[Validación principal]\n"
  run_category_linux "$cid" || { printf "FAIL validación principal\n"; return 1; }

  printf "\n[Validaciones derivadas de metadatos]\n"
  [[ "$multi" == "Si" ]] && run_multiline_local
  [[ "$dl" == "Si" ]] && [[ "$cid" != "12" ]] && run_local_http_fetch
  [[ "$b64" == "Si" ]] && run_base64_roundtrip
  [[ "$obf" == "Si" ]] && run_obfuscation_roundtrip
  [[ "$hid" == "Si" ]] && run_quiet_local
  if [[ "$pers" == "Si" ]]; then
    printf "INFO persistencia=Si se conserva como metadato; el launcher no modifica mecanismos de inicio/persistencia del sistema.\n"
  fi
  printf "\nResultado: flujo finalizado.\n"
}

browse_category() {
  local cid="$1"
  local -a groups=()
  local grp_line
  while IFS= read -r grp_line; do groups+=("$grp_line"); done < <(groups_for_category "$cid")
  show_category_detail "$cid"
  if (( ${#groups[@]} == 0 )); then
    printf "\nNo hay grupos estructurales para esta categoría.\n"
    pause_screen
    return
  fi

  local page=0 pages choice start end idx line grp rep nreg dl b64 obf hid
  pages=$(( (${#groups[@]} + PAGE_SIZE - 1) / PAGE_SIZE ))
  while true; do
    clear_screen
    print_banner
    show_category_detail "$cid"
    start=$((page * PAGE_SIZE))
    end=$((start + PAGE_SIZE))
    (( end > ${#groups[@]} )) && end=${#groups[@]}
    printf "\nGrupos %d-%d de %d (página %d/%d)\n\n" "$((start+1))" "$end" "${#groups[@]}" "$((page+1))" "$pages"
    for ((idx=start; idx<end; idx++)); do
      line="${groups[$idx]}"
      tsv_split "$line" _ _ _ grp rep nreg _ _ _ _ _ _ _ _ _ dl b64 obf hid _ _ _
      printf "%4d) grupo=%-5s rep=%-5s reg=%-5s flags[D:%s B:%s O:%s H:%s]\n" \
        "$((idx+1))" "$grp" "$rep" "$nreg" "$dl" "$b64" "$obf" "$hid"
    done
    printf "\nNúmero=abrir | n=siguiente | p=anterior | b=volver\n> "
    IFS= read -r choice || return
    case "$(lc "$choice")" in
      b) return ;;
      n) (( page+1 < pages )) && page=$((page+1)) ;;
      p) (( page > 0 )) && page=$((page-1)) ;;
      *)
        if [[ "$choice" =~ ^[0-9]+$ ]] && (( choice>=1 && choice<=${#groups[@]} )); then
          line="${groups[$((choice-1))]}"
          clear_screen
          print_banner
          show_group_detail_line "$line"
          printf "\n1) Ejecutar validación local\n2) Volver\n> "
          IFS= read -r action
          if [[ "$action" == "1" ]]; then
            run_group_safe "$line"
            pause_screen
          fi
        fi
        ;;
    esac
  done
}

browse_categories() {
  local cid
  while true; do
    clear_screen
    print_banner
    list_categories
    printf "\nEscribe el ID de categoría (p. ej. 08, 11, ZZ) o b para volver:\n> "
    IFS= read -r cid || return
    [[ "$(lc "$cid")" == "b" ]] && return
    if category_line_by_id "$cid" >/dev/null && [[ -n "$(category_line_by_id "$cid")" ]]; then
      browse_category "$cid"
    else
      printf "ID no válido.\n"
      sleep 1
    fi
  done
}

search_group() {
  local q line count
  printf "Grupo numérico o fragmento de hash: "
  IFS= read -r q
  [[ -z "$q" ]] && return
  count=0
  while IFS= read -r line; do
    [[ -z "$line" ]] && continue
    count=$((count+1))
    show_group_detail_line "$line"
    if (( count >= 20 )); then
      printf "\nSe muestran como máximo 20 coincidencias.\n"
      break
    fi
  done < <(printf '%s\n' "$GROUP_TSV" | awk -F '\t' -v q="$q" 'NR>1 && ($4==q || index(tolower($22),tolower(q))>0)')
  (( count == 0 )) && printf "Sin coincidencias.\n"
  pause_screen
}

show_metadata() {
  clear_screen
  print_banner
  printf '%s\n' "$META_TSV" | awk -F '\t' 'NR>1 {printf "%-18s | %-34s | %s\n",$1,$2,$3}'
  pause_screen
}

self_test() {
  printf "Autoprueba del launcher (sin red externa)\n"
  local fail=0
  [[ "$(printf '%s\n' "$SUMMARY_TSV" | awk 'END{print NR-1}')" == "16" ]] || fail=1
  [[ "$(printf '%s\n' "$GROUP_TSV" | awk 'END{print NR-1}')" == "1483" ]] || fail=1
  printf "OK  taxonomía integrada: 16 categorías / 1483 grupos\n"
  run_multiline_local || fail=1
  run_base64_roundtrip || fail=1
  run_obfuscation_roundtrip || fail=1
  run_quiet_local || fail=1
  run_local_http_fetch || fail=1
  if (( fail==0 )); then
    printf "SELF-TEST: OK\n"
    return 0
  else
    printf "SELF-TEST: CON FALLOS\n"
    return 1
  fi
}

main_menu() {
  local opt
  while true; do
    clear_screen
    print_banner
    printf "1) Explorar categorías y grupos\n"
    printf "2) Buscar grupo por ID/hash\n"
    printf "3) Ver metadatos del dataset\n"
    printf "4) Ejecutar autoprueba local\n"
    printf "0) Salir\n\n> "
    IFS= read -r opt || exit 0
    case "$opt" in
      1) browse_categories ;;
      2) search_group ;;
      3) show_metadata ;;
      4) clear_screen; print_banner; self_test; pause_screen ;;
      0) exit 0 ;;
      *) printf "Opción no válida.\n"; sleep 1 ;;
    esac
  done
}

case "${1:-}" in
  --self-test) self_test ;;
  --list) list_categories ;;
  --help|-h)
    printf "Uso: %s [--self-test|--list]\n" "$SCRIPT_NAME"
    ;;
  *) main_menu ;;
esac
