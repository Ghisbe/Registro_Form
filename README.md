# Registro_Form

🗂 Proyecto: Registro de Clientes con Formularios VBA en Excel

- Objetivo: Diseñar un sistema completo de registro de clientes que permita agregar, editar, eliminar, buscar y validar datos utilizando formularios personalizados y VBA en Excel.

Este proyecto es parte del curso de : https://eltiotech.com/curso-completo-vba-macros-excel/

![image.png](attachment:c3f2a905-c3cc-4c7b-9669-1002ba667e3e:image.png)

✅ Paso a Paso:
1. Diseño del Formulario Principal
Se crea el formulario UserForm con los campos: Nombre, Apellido, Fecha de nacimiento, Edad, Email, Teléfono, País, Provincia/Estado y Ciudad.

Se añaden controles: TextBox, ComboBox, CommandButton, Calendar, etc.

Botones: Registrar, Modificar, Limpiar.

![image.png](attachment:27012ab1-f2d0-4ec3-9f43-34a77589e28e:image.png)

2. Validación de Datos
Se incorporan rutinas de validación:

Solo texto en nombres/apellidos.

Solo números en campos numéricos.

Estructura de email válida.

Campos obligatorios no vacíos.

3. Insertar Calendario.
Se integra un formulario de calendario (frmCalendario) para seleccionar fecha de nacimiento.

4. Programar el Botón "Registrar"
Se programa para guardar los datos ingresados en la hoja BD.

Se asigna un código único tipo CLI_0001.

5. Eliminar Registros
Se crea un formulario de búsqueda con ListBox.

![image.png](attachment:05d10dfc-b57b-4471-8c56-9a7a60cf2fec:image.png)

Al seleccionar un registro y confirmar con una contraseña (TextBox1 = "12345") en este formulario, el sistema elimina toda la fila correspondiente.

![image.png](attachment:dcee0cc0-afc5-43cc-9205-cb50aecd51ec:image.png)

6. Generar Código Automático
Cada nuevo registro obtiene un código consecutivo.

Se verifica el último código existente en la hoja y se incrementa.

7. Botón Buscar y Modificar
Se permite la búsqueda múltiple por nombre, apellido o código.

Al hacer clic, se cargan los datos en el formulario para editar y guardar.

Se muestra al hacer doble clic.

8. Botón Limpiar Campos
Borra todos los campos del formulario de forma automática.

9. Protección de Datos y Fluidez del Usuario
Se usa Unload Me para cerrar formularios.

Mensajes de confirmación (MsgBox) para acciones críticas.

Uso de módulos estándar y formularios bien estructurados.

### Resultado Final

El resultado es un **formulario completo y funcional en Excel**, que permite registrar, modificar, eliminar y buscar clientes de manera organizada. Se incluye validación automática de datos, control de errores, cálculo de edad, y una interfaz amigable, ideal para gestionar bases de datos pequeñas y medianas dentro de un entorno de oficina.

### Próximos pasos

Como próximos pasos, voy a seguir profundizando en el uso de formularios en Excel y Visual Basic, explorando nuevas funciones como la conexión con bases de datos, automatización avanzada y mejoras en la interfaz de usuario. Excel es una herramienta poderosa que, con cada nueva función que descubro, demuestra tener un potencial sorprendente para optimizar tareas y resolver problemas de forma creativa.

