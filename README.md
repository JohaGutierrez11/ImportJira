# Antecedentes

En los procesos de aseguramiento de la calidad, los casos de prueba suelen gestionarse mediante hojas de cálculo en formato Excel, utilizando lenguaje Gherkin para su posterior automatización. Sin embargo, el template utilizado para la definición de estos casos de prueba no es compatible directamente con el formato requerido por Jira para su carga e integración.

Esta incompatibilidad genera reprocesos y errores manuales, lo que motiva el desarrollo de una aplicación orientada a la transformación automática de los templates de casos de prueba, permitiendo su adaptación al formato requerido por Jira de manera eficiente y controlada.

# ExcelSuiteReader

Herramienta Java para la lectura y procesamiento de archivos Excel, empaquetada como **fat JAR** para su ejecución directa sin dependencias externas.

---

## ✅ Requisitos

* **Java JDK 17** (obligatorio)
* Sistema operativo: Windows / Linux / macOS

> ⚠️ Importante: el proyecto fue **compilado con Java 17**. Versiones anteriores (Java 8 u 11) provocarán errores al ejecutar el JAR.

---

## 🔍 Verificar versión de Java instalada

Antes de ejecutar el JAR, valida la versión de Java:

```bash
java -version
```

Salida esperada (ejemplo):

```
java version "17.x.x"
```

Si la versión es menor a 17, debes instalar JDK 17 y configurar la variable de entorno `JAVA_HOME`.

---

## 📦 Construcción del JAR (opcional)

Si deseas generar el JAR desde el código fuente:

```bash
gradle clean shadowJar
```

El archivo se generará en:

```
build/libs/
```

Con un nombre similar a:

```
excel-suite-reader-2.0-all.jar
```

---

## ▶️ Ejecución del JAR

Desde una terminal, ubícate en el directorio donde se encuentra el JAR y ejecuta:

```bash
java -jar excel-suite-reader-2.0-all.jar
```

> ✅ **Nota**: utiliza siempre el archivo que termina en `-all.jar`, ya que incluye todas las dependencias necesarias.

---

## ❌ Errores comunes

### Error: `A JNI error has occurred`

Causa probable:

* El JAR está siendo ejecutado con una versión de Java distinta a Java 17.

Solución:

* Verificar `java -version`
* Asegurar que Java 17 esté configurado en el `PATH`

---

## ℹ️ Información técnica

* **Lenguaje**: Java
* **Versión de Java utilizada**: **Java 17**
* **Build tool**: Gradle
* **Empaquetado**: Fat JAR mediante Shadow Plugin

---

## 👤 Autor

Johanna Gutiérrez

Fecha de generación: 2025

---
