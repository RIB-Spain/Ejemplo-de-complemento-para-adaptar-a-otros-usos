# Complemento para Presto: Convertir resúmenes a minúsculas

Complemento VBS para Presto que convierte los resúmenes de las partidas (conceptos de naturaleza 5) a minúsculas, respetando dos reglas: la letra inicial del resumen se mantiene en mayúscula, y las palabras que contienen números no se modifican.

## Funcionamiento

- Recorre todos los conceptos de la obra activa filtrados por partidas.
- Convierte el resumen a minúsculas salvo la inicial y los términos alfanuméricos (p. ej. `B500S`, `16mm`).
- Los cambios son reversibles mediante la función Deshacer de Presto.
- Registra en el log los valores originales y modificados para su revisión.

## Uso

1. Abrir la obra en Presto.
2. Ejecutar el complemento desde el menú de complementos de Presto.
3. Confirmar la ejecución en el cuadro de diálogo.
4. Revisar el log para comprobar los cambios realizados.

## Notas

- Las palabras que contienen al menos un dígito (p. ej. `HA-25`, `IPE200`) no se convierten a minúsculas.
- El complemento activa automáticamente el mecanismo de Deshacer de Presto, por lo que los cambios pueden revertirse de forma segura.
- El núcleo de las funciones auxiliares fue generado originalmente con OpenAI GPT-3.5.
