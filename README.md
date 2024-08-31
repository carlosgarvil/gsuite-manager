# Google Apps Script for Google Workspace Management

Este repositorio contiene un conjunto de funciones de Google Apps Script diseñadas para gestionar usuarios y grupos en Google Workspace. Este script surge como alternativa personalizada a la extensión "Ok Goldy", que actualmente no funciona (agosto 2024). Las funciones permiten crear usuarios, añadirlos a grupos, eliminar usuarios, suspender cuentas, y otras operaciones administrativas de manera eficiente.

## Funciones Principales

### 1. `createUser()`

**Descripción**: Crea un usuario en Google Workspace con los datos especificados en una hoja de cálculo y lo asigna a una unidad organizativa específica. Si el correo pertenece a un alumno, también se añade al grupo especificado.

- **Columnas de Entrada**:
  - **Columna A**: Nombre Gsuite
  - **Columna B**: Apellidos Gsuite
  - **Columna C**: Email
  - **Columna D**: Email Grupo
  - **Columna E**: Contraseña

- **Columnas de Salida**:
  - **Columna F**: Estado de creación del usuario (`Created` o mensaje de error).
  - **Columna G**: Unidad organizativa asignada.
  - **Columna H**: Estado de la asignación al grupo (`Added to Group` o mensaje de error).

### 2. `addToGroup()`

**Descripción**: Añade usuarios a los grupos especificados en una hoja de cálculo.

- **Columnas de Entrada**:
  - **Columna A**: Email del usuario
  - **Columna B**: Email del grupo

- **Columnas de Salida**:
  - **Columna C**: Resultado (`Added` o mensaje de error).

### 3. `emptyGroup()`

**Descripción**: Elimina todos los usuarios de los grupos especificados en una hoja de cálculo.

- **Columna de Entrada**:
  - **Columna A**: Lista de grupos (desde la celda A2 en adelante).

### 4. `copyGroups()`

**Descripción**: Copia la pertenencia a grupos de un usuario a otro.

- **Celda A2**: Email del usuario 1 (origen).
- **Celda B2**: Email del usuario 2 (destino).
- **Columna C**: Lista de grupos del usuario 1.
- **Columna D**: Resultado de la asignación del usuario 2 a los grupos (o mensaje de error).

### 5. `suspendUsers()`

**Descripción**: Suspende usuarios listados en una hoja de cálculo.

- **Columna de Entrada**:
  - **Columna A**: Lista de correos de los usuarios a suspender.

- **Columna de Salida**:
  - **Columna B**: Resultado de la suspensión (`Suspended` o mensaje de error).

- **Celda E1**: Número de usuarios procesados, actualizado en intervalos de 5.

### 6. `getTeachers()`

**Descripción**: Obtiene el nombre, apellidos y email de los usuarios del grupo "Claustro".

- **Columnas de Salida**:
  - **Columna A**: Nombre
  - **Columna B**: Apellidos
  - **Columna C**: Email

- **Celda F1**: Tiempo total de ejecución del script en segundos.

### 7. `createGroup()`

**Descripción**: Crea un nuevo grupo en Google Workspace con los detalles especificados en una hoja de cálculo.

- **Columnas de Entrada**:
  - **Columna A**: Nombre del grupo
  - **Columna B**: Email del grupo
  - **Columna C**: Descripción

### 8. `deleteGroups()`

**Descripción**: Elimina grupos especificados en una hoja de cálculo.

- **Columna de Entrada**:
  - **Columna A**: Lista de grupos a eliminar.

## Cómo Usar el Script

1. **Configuración Inicial**: Añade las funciones a un proyecto de Google Apps Script vinculado a una hoja de cálculo de Google Sheets.
2. **Asignar Permisos**: Asegúrate de que el proyecto tenga permisos para acceder y modificar usuarios y grupos mediante la API de Admin SDK.
3. **Ejecutar Funciones**: Puedes ejecutar las funciones desde el editor de Apps Script o creando botones en la hoja de cálculo para su ejecución.

## Notas Importantes

- **Permisos de Admin SDK**: Estas funciones requieren que tengas los permisos adecuados como administrador en Google Workspace.
- **Manejo de Errores**: Los errores durante la ejecución se registran en la hoja de cálculo correspondiente, proporcionando detalles para la depuración.
- **Optimización**: Se han implementado técnicas para minimizar el impacto en el rendimiento, como el uso de `SpreadsheetApp.flush()` y la actualización de datos en bloques.

## Contribuir

Si deseas mejorar o añadir nuevas funciones, por favor, abre un pull request con tus cambios y una descripción detallada de lo que has modificado.

## Licencia

Este proyecto está licenciado bajo la Licencia Pública General de GNU versión 3 (GPL-3.0).
