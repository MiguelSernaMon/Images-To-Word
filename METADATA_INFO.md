# 📊 Extracción de Metadata y Fecha de Envío

## ¿Qué información se extrae?

La aplicación extrae automáticamente la **fecha y hora** de cada imagen para:
1. **Ordenar** las imágenes cronológicamente
2. **Mostrar** la fecha encima de cada foto en el documento Word

### Fecha y Hora 📅
La aplicación busca la fecha/hora en este orden de prioridad:

1. **Nombre del archivo WhatsApp**: Formato `IMG-YYYYMMDD-WA####.jpg`
   - Ejemplo: `IMG-20231225-WA0001.jpg` → 25 de diciembre de 2023

2. **EXIF DateTimeOriginal**: Fecha en que se tomó la foto originalmente

3. **EXIF DateTime**: Fecha de última modificación registrada en EXIF

4. **Fecha de modificación del archivo**: Como último recurso

## Cómo funciona el ordenamiento por fecha

Cuando seleccionas **"Fecha de envío"** como método de ordenamiento:

- Las imágenes se ordenan **cronológicamente** (de más antigua a más reciente)
- **Cada imagen muestra** su fecha y hora encima en el documento Word
- Formato mostrado: `📅 DD/MM/YYYY 🕐 HH:MM:SS`

### Ejemplo visual en el Word:
```
┌─────────────────────────┐
│  📅 25/12/2023 � 10:30:15 │
│                         │
│      [IMAGEN 1]         │
│                         │
└─────────────────────────┘

┌─────────────────────────┐
│  📅 25/12/2023 🕐 15:45:22 │
│                         │
│      [IMAGEN 2]         │
│                         │
└─────────────────────────┘
```

## Limitaciones de WhatsApp

WhatsApp modifica las imágenes al enviarlas:

- ❌ **Elimina** la mayoría de metadata EXIF (GPS, autor, cámara, etc.)
- ❌ **Comprime** las imágenes (reduce calidad)
- ✅ **Mantiene** el nombre de archivo con fecha (formato IMG-YYYYMMDD)
- ✅ **Mantiene** la fecha de modificación del archivo

⚠️ **Importante**: La información del **remitente NO está disponible** en las imágenes de WhatsApp por razones de privacidad.

## Botón "Analizar Metadata" 🔍

Úsalo **antes** de convertir para:
- Ver qué fecha se detectó para cada imagen
- Verificar si el ordenamiento será correcto
- Identificar imágenes sin fecha detectada
