/* DendroRed — configuración
 * Copia estos dos valores desde Supabase: Project Settings > API (o "API Keys").
 *   supabaseUrl:   Project URL, p. ej. https://abcdefghijkl.supabase.co
 *   supabaseClave: la clave pública ("publishable" o "anon"). Es segura en la página:
 *                  lo que cada persona puede hacer lo deciden las reglas de la base.
 * NUNCA pongas aquí la clave secreta ("secret" o "service_role").
 */
window.DENDRORED_CONFIG = {
  supabaseUrl: 'https://xyuwihzanrkxzepslmbi.supabase.co',
  supabaseClave: 'sb_publishable_-TrWb2Pz08_2pgO9UvZ5SQ_9Fm8r3GE',

  // Encuadre inicial del mapa (latitud, longitud) y acercamiento
  centroMapa: [-39.5, -72.0],
  zoomMapa: 5,

  // Clave de acceso de Web3Forms, para el formulario de la pestaña "Contacto".
  // Se obtiene gratis en web3forms.com indicando la dirección que recibirá los
  // mensajes. Es pública: solo sirve para enviar a esa dirección, no la revela.
  web3formsClave: '83d72e1f-60eb-440a-b114-d61ec5e7801c',

  // Correo(s) visibles en "Acerca de" (opcional). Déjalo vacío para que las
  // consultas pasen solo por el formulario y no se exponga ninguna dirección.
  // Acepta un correo suelto o una lista.
  correoContacto: ''
};
