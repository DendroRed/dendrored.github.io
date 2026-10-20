# DendroRed

Catálogo colaborativo de sitios y cronologías dendrocronológicas de Chile.
La página se aloja gratis en GitHub Pages y los datos viven en Supabase.

- Cualquier persona ve el mapa, filtra, abre las fichas y exporta (Excel, CSV o GeoJSON).
- Para registrar o editar hay que crear una cuenta, que un administrador aprueba.
- Cada investigador edita solo lo suyo y puede agregar su cronología a un sitio de otra persona.
- Cada cronología distingue a su **responsable**, que responde por ella, de sus **autores**, que la cofecharon. Puede ser la misma persona o no.
- La serie de tiempo no es obligatoria. Si está "a pedido", los usuarios aprobados ven el correo del responsable.
- Los correos de los usuarios nunca son públicos.

## Archivos y carpetas

Respeta esta estructura al subir los archivos a GitHub:

```
dendrored/
├── index.html
├── LEEME.md
├── plantilla_importacion.csv
├── css/
│   └── estilos.css
├── js/
│   ├── config.js        ← aquí van la dirección y la clave de Supabase
│   ├── app.js
│   ├── coordenadas.js
│   └── importador.js
├── img/
│   └── logo.svg
└── .github/
    └── workflows/
        └── mantener-activo.yml
```

El archivo `dendrored_esquema.sql` no va en el repositorio: se ejecuta una vez en Supabase.

## 1. Crear la base de datos en Supabase

1. Crea una cuenta en supabase.com. Para que después sea fácil traspasar el proyecto, crea primero una **organización** llamada DendroRed y el proyecto dentro de ella. Como región, la más cercana es São Paulo.
2. En el proyecto, abre **SQL Editor**, pega el contenido completo de `dendrored_esquema.sql` y presiona **Run**.
   - Si ya habías ejecutado la versión 0.1, primero ejecuta este bloque para borrarla (no hay datos que perder todavía):
     ```sql
     drop trigger if exists al_registrarse on auth.users;
     drop view if exists public.catalogo;
     drop table if exists public.historial, public.cronologia_autores, public.cronologias, public.sitios,
       public.importaciones, public.especies, public.perfiles, public.laboratorios cascade;
     drop function if exists public.es_aprobado, public.es_curador, public.es_admin, public.crear_perfil,
       public.proteger_perfil, public.marcar_actualizacion, public.fijar_autoria, public.registrar_historial,
       public.sitios_cercanos, public.correo_de, public.administrar_cuentas cascade;
     drop type if exists public.rol_usuario, public.disponibilidad_serie, public.precision_coord,
       public.parametro_crono, public.estado_crono, public.formato_archivo cascade;
     ```
3. En **Project Settings > API** (o **API Keys**) copia la **Project URL** y la clave **pública** ("publishable" o "anon"). Escríbelas en `js/config.js`.
   **Nunca** pongas en la página la clave secreta ("secret" o "service_role").

## 2. Publicar la página en GitHub Pages

1. En GitHub crea una **organización** llamada DendroRed (en vez de usar tu cuenta personal, para poder sumar administradores después).
2. Crea un repositorio público y sube todos los archivos respetando las carpetas.
   - Si el repositorio se llama `dendrored`, la dirección será `https://dendrored.github.io/dendrored/`.
   - Si se llama `dendrored.github.io`, la dirección será más corta: `https://dendrored.github.io/`.
   - El nombre de la organización puede estar tomado; en ese caso usa otro y cambia la dirección en consecuencia.
3. En el repositorio: **Settings > Pages > Build and deployment**, elige **Deploy from a branch**, rama `main`, carpeta `/ (root)`, y guarda. En uno o dos minutos la página queda publicada.

## 3. Conectar las cuentas de usuario

En Supabase, **Authentication > URL Configuration**:

- **Site URL**: la dirección de tu página (por ejemplo `https://dendrored.github.io/dendrored/`).
- **Redirect URLs**: agrega la misma dirección. Sin esto, los enlaces de confirmación y de cambio de contraseña no vuelven a la página.

Supabase pide confirmar el correo al registrarse. El servicio de correo que trae de fábrica tiene un límite bajo de envíos por hora, suficiente para empezar. Si se registran muchas personas a la vez, configura un servidor de correo propio en **Authentication > Emails > SMTP Settings**.

## 4. Hacerte administrador

1. Abre la página, entra a **Ingresar > Crear cuenta** y regístrate. Confirma tu correo.
2. En el **SQL Editor** de Supabase ejecuta, con tu correo:
   ```sql
   update public.perfiles
      set rol = 'administrador', aprobado = true
    where correo = 'tu-correo@ejemplo.cl';
   ```
3. Recarga la página. Aparecerá **Administrar**, donde puedes:
   - aprobar cuentas y asignar roles (investigador, curador, administrador),
   - agregar laboratorios (los usuarios eligen el suyo en su perfil),
   - corregir especies y sus códigos ITRDB,
   - deshacer una importación completa (ver más abajo),
   - revisar el historial de cambios.

Para traspasar la administración más adelante, nombra administrador a otra persona desde **Administrar > Cuentas** y agrégala como miembro de las organizaciones de GitHub y Supabase.

## 5. Mantener activa la base gratuita

Los proyectos gratuitos de Supabase se pausan tras una semana sin uso. La tarea `mantener-activo.yml` hace una consulta liviana cada tres días. Para activarla, en el repositorio de GitHub: **Settings > Secrets and variables > Actions > New repository secret**, y crea dos:

- `SUPABASE_URL`: la misma dirección de `js/config.js`
- `SUPABASE_CLAVE`: la misma clave pública de `js/config.js`

Puedes probarla a mano desde la pestaña **Actions > Mantener activa la base de datos > Run workflow**.

## Importar datos

En **Importar** se aceptan CSV o TXT (separados por comas, punto y coma o tabulaciones), Excel, GPX, KML y GeoJSON.

- Cada fila es un sitio, o un sitio con una cronología. Un sitio con varias cronologías repite su código en varias filas.
- Si un código ya existe en el catálogo, solo se agregan las cronologías a ese sitio.
- Las columnas se asignan solas cuando sus nombres son reconocibles; si no, se eligen a mano.
- Coordenadas: grados decimales, grados-minutos-segundos o UTM (zonas 17 a 20 sur), en WGS84, PSAD56 o SAD69. Todo se guarda en WGS84.
- Antes de guardar se muestran las filas con errores (que no se guardan) y avisos, como sitios a menos de 500 m de otro ya registrado.
- La `plantilla_importacion.csv` trae filas de ejemplo con todos los campos.

### Autoría al importar

Cuatro columnas deciden a quién se atribuye cada registro. Son opcionales, pero conviene usarlas cuando el archivo incluye trabajo de otras personas:

| Columna | Qué guarda |
|---|---|
| `laboratorio` | El laboratorio que muestreó el **sitio** |
| `laboratorio_crono` | El laboratorio que desarrolló la **cronología**, si es otro |
| `responsable` | Quien responde por la cronología |
| `autores` | Quienes la hicieron, **separados por punto y coma** |

**Responsable y autor no son lo mismo.** El responsable responde por la cronología; el autor es quien la hizo, quien cofechó las muestras. A veces es la misma persona y a veces no, y esa diferencia es un dato:

- Si el responsable **también cofechó**, ponlo en las dos columnas. Queda como responsable y como autor, sin repetirse en la ficha.
- Si el responsable **solo responde**, no lo pongas en `autores`. La ficha lo mostrará como responsable y nada más.
- La página no deduce una cosa de la otra. Si dejas `autores` vacía, la cronología queda sin autores, no con el responsable puesto ahí.
- Si pones `autores` y dejas `responsable` vacía, quedan solo como autores: nadie pasa a ser responsable por estar primero.

Y además:

- Los nombres de laboratorio deben coincidir con los de **Administrar → Laboratorios** (no distingue tildes ni mayúsculas, y también acepta el nombre de la institución). Los que no se reconozcan se avisan al terminar y el registro queda sin laboratorio.
- Si el nombre de una persona coincide con el de una cuenta registrada, se vincula a su perfil; si no, se guarda como texto.
- El orden del archivo es el orden en que se citan.
- **Si el archivo no trae ninguna de estas columnas, quien importa queda como responsable de todas las cronologías (no como autor) y su laboratorio se asigna a todos los sitios.** El resumen final lo advierte.

El filtro de laboratorio del mapa busca en los dos niveles: un sitio aparece si el laboratorio elegido muestreó el sitio **o** desarrolló alguna de sus cronologías.

### Deshacer una importación

En **Administrar → Importaciones** está la lista de archivos subidos, con cuántos sitios y cuántas cronologías sigue teniendo cada uno. **Borrar datos** elimina todo lo que creó esa importación. Sirve cuando un archivo entró con las columnas corridas o con la autoría equivocada: es más rápido corregir el archivo y volver a subirlo que editar decenas de fichas.

- Las cronologías se borran primero, porque algunas pueden estar en sitios que ya existían y que no hay que tocar.
- Borrar un sitio arrastra **todas** sus cronologías, incluidas las que otra persona le agregó después. Si eso va a pasar, la tabla lo avisa antes (*"+N cronologías ajenas"*) y la confirmación lo repite.
- El registro de la importación queda en la lista, con 0 sitios y 0 cronologías, como constancia de que se hizo.
- No se puede deshacer.

## Coordenadas sensibles

En cada sitio, o en una importación completa, se puede marcar **Publicar coordenadas aproximadas**. Se redondean a dos decimales (~1 km) y no se guarda la coordenada original. Úsalo para especies amenazadas o sitios vulnerables a la extracción.

## Formulario de contacto

La pestaña **Contacto** deja que cualquier persona escriba sin registrarse y sin que nadie publique su dirección de correo. Los mensajes se envían con [Web3Forms](https://web3forms.com), que es gratuito:

1. En web3forms.com indica la dirección que recibirá los mensajes. Te llega una **clave de acceso** por correo.
2. Escríbela en `js/config.js`, en `web3formsClave`. Es una clave pública: solo sirve para enviar mensajes a esa dirección, no la revela ni permite leer nada.
3. Al crear el formulario, Web3Forms pide la **dirección del sitio**. Mientras pruebas en tu computador va `localhost`. **Cuando publiques la página en GitHub Pages, vuelve a Web3Forms y cámbiala por la dirección definitiva**, o los mensajes podrían dejar de llegar sin aviso.

El plan gratuito admite **un solo destinatario**. Si los mensajes deben llegar a varias personas, crea un grupo de correo (por ejemplo un Grupo de Google) con todas ellas y registra la dirección del grupo. Así se suman o quitan destinatarios sin tocar el código.

Mientras `web3formsClave` esté vacía, la pestaña avisa que falta configurarla. El formulario descarta los envíos de robots con una casilla oculta.

Si además quieres mostrar una dirección de correo en **Acerca de**, llena `correoContacto` en `js/config.js` (acepta un correo o una lista). Déjalo vacío para que todo pase por el formulario.

## Probar la página en tu computador

Abrir `index.html` con doble clic no basta, porque el inicio de sesión necesita una dirección web. Desde la carpeta del proyecto ejecuta:

```
python -m http.server 8000
```

y abre `http://localhost:8000`. Agrega esa dirección también en **Redirect URLs** de Supabase si quieres probar el registro de cuentas.
