# -*- coding: utf-8 -*-
"""
Diccionario de motivos de rechazo (canon en ingles del gateway)
con traduccion al espanol, explicacion y accion recomendada.

Cada entrada:
    "clave canonica en minusculas (substring suficiente para match)": {
        "es": "Texto en espanol",
        "explica": "Que esta pasando",
        "accion": "Que debe hacerse",
        "responsable": "Cliente | Banco | Operacion | Pasarela | Comercial",
    }

El matching se hace por substring case-insensitive sobre el motivo original.
La primera entrada que coincide es la que se usa.
"""

MOTIVO_CATALOG = [
    {
        "match": "decline - general decline of the card",
        "es": "Rechazo general del banco emisor",
        "explica": "El banco rechazo el cobro sin dar razon especifica. Suele indicar fondos insuficientes intermitentes, cliente bloqueado por fraude del banco, o limites internos del emisor.",
        "accion": "Contactar al socio para que valide su tarjeta con el banco. Reintentar 24-72h. Ofrecer cambiar a otro medio de pago.",
        "responsable": "Cliente",
    },
    {
        "match": "not authorised - low funds",
        "es": "Sin saldo / fondos insuficientes",
        "explica": "La tarjeta del cliente no tiene saldo disponible al momento del cobro.",
        "accion": "Reintentar al final del mes o cuando se acredite la nomina del cliente. Notificar al socio para que recargue.",
        "responsable": "Cliente",
    },
    {
        "match": "not authorised - do not honor",
        "es": "Banco no autoriza ('do not honor')",
        "explica": "El banco rechaza la transaccion por reglas internas (sospecha, perfil de riesgo, transaccion fuera de patron).",
        "accion": "Pedir al cliente que llame al banco a autorizar pagos recurrentes EVO. Reintentar despues de la gestion.",
        "responsable": "Cliente",
    },
    {
        "match": "no incluye bloqueadas o timeout",
        "es": "Tarjeta bloqueada o timeout en la red",
        "explica": "La tarjeta esta bloqueada (por el cliente o el banco) o la red de la pasarela no respondio a tiempo.",
        "accion": "Reintentar mas tarde si fue timeout. Si el bloqueo persiste, contactar al cliente para actualizar el medio de pago.",
        "responsable": "Pasarela",
    },
    {
        "match": "tarjeta de uso restringido, no permite retiros",
        "es": "Tarjeta restringida (no permite cobros recurrentes)",
        "explica": "El plastico es debito de uso restringido y el banco no permite cargos automaticos.",
        "accion": "Solicitar al cliente que registre una tarjeta de credito o debito habilitada para recurrencias.",
        "responsable": "Cliente",
    },
    {
        "match": "tarjeta no registrada",
        "es": "Tarjeta no registrada en el sistema",
        "explica": "No hay tarjeta asociada al socio en EVO al momento del cobro. Falta el alta del medio de pago.",
        "accion": "Operacion: completar el alta del medio de pago en recepcion. Cruzar contra base de afiliacion.",
        "responsable": "Operacion",
    },
    {
        "match": "not authorised - restricted",
        "es": "Tarjeta con uso restringido por el banco",
        "explica": "El banco emisor marca la tarjeta como restringida para este tipo de cobro.",
        "accion": "Solicitar al cliente que el banco habilite cobros recurrentes o cambiar de medio de pago.",
        "responsable": "Cliente",
    },
    {
        "match": "life cycle. (mastercard use only)",
        "es": "Tarjeta vencida / reemplazada (Mastercard)",
        "explica": "La tarjeta del socio fue reemplazada o cancelada por Mastercard (cambio de plastico).",
        "accion": "Pedir al cliente actualizar la tarjeta. Considerar Account Updater de Mastercard si esta disponible.",
        "responsable": "Cliente",
    },
    {
        "match": "operación mal sucedida",
        "es": "Operacion fallida en la pasarela",
        "explica": "Error generico devuelto por la pasarela; suele ser problema de red o respuesta invalida del banco.",
        "accion": "Reintentar de inmediato; si persiste, escalar a soporte de pasarela.",
        "responsable": "Pasarela",
    },
    {
        "match": "unable to authorise - no account",
        "es": "Cuenta no encontrada en el banco",
        "explica": "El emisor responde que la cuenta asociada a la tarjeta no existe.",
        "accion": "Contactar al socio para verificar y actualizar el medio de pago.",
        "responsable": "Cliente",
    },
    {
        "match": "not authorised - not permitted to terminal",
        "es": "Terminal no autorizada por el banco",
        "explica": "El banco no permite pagos desde este tipo de comercio o terminal.",
        "accion": "Soporte de pasarela: revisar configuracion MCC del comercio EVO. Contactar al banco emisor si es masivo.",
        "responsable": "Pasarela",
    },
    {
        "match": "one or more fields in the request contains invalid data",
        "es": "Datos invalidos enviados a la pasarela",
        "explica": "La pasarela detecta campos malformados (numero, fecha, CVV, monto fuera de rango).",
        "accion": "Operacion: revisar la integracion. Validar que la captura de datos en recepcion sea correcta.",
        "responsable": "Operacion",
    },
    {
        "match": "profile not found",
        "es": "Perfil del socio no encontrado",
        "explica": "El registro del cliente o del medio de pago no existe en la pasarela al momento del cobro.",
        "accion": "Operacion: revisar el alta del cliente; sincronizar perfiles entre EVO y pasarela.",
        "responsable": "Operacion",
    },
    {
        "match": "incorrect pin",
        "es": "PIN incorrecto",
        "explica": "El cobro requirio PIN y el ingresado no es valido.",
        "accion": "Cobros automaticos no deberian pedir PIN. Revisar configuracion del medio de pago con la pasarela.",
        "responsable": "Pasarela",
    },
    {
        "match": "token temporarily suspended",
        "es": "Token suspendido (requiere actualizacion)",
        "explica": "El token tokenizado de la tarjeta fue suspendido y necesita refresh.",
        "accion": "Activar Account Updater. Pedir al cliente actualizar la tarjeta.",
        "responsable": "Cliente",
    },
    {
        "match": "amount too large",
        "es": "Monto excede el limite del banco",
        "explica": "El cobro supera el tope autorizado por el banco emisor para esta tarjeta.",
        "accion": "Dividir el cobro o solicitar al cliente que aumente el cupo con su banco.",
        "responsable": "Cliente",
    },
    {
        "match": "closed account",
        "es": "Cuenta cerrada",
        "explica": "El cliente cerro la cuenta o el banco la dio de baja.",
        "accion": "Imposible recuperar via reintento. Contactar al socio para nuevo medio de pago.",
        "responsable": "Cliente",
    },
    {
        "match": "invalid card number",
        "es": "Numero de tarjeta invalido",
        "explica": "Los digitos almacenados de la tarjeta no son validos (digitacion mala o tarjeta vencida).",
        "accion": "Operacion: pedir al cliente reingresar la tarjeta en recepcion.",
        "responsable": "Operacion",
    },
    {
        "match": "tarjeta robada o extraviada",
        "es": "Tarjeta robada o extraviada",
        "explica": "El cliente reporto perdida o robo; el banco la bloqueo.",
        "accion": "Imposible recuperar. Contactar al cliente para nuevo medio de pago.",
        "responsable": "Cliente",
    },
    {
        "match": "invalid transaction",
        "es": "Transaccion invalida",
        "explica": "El banco considera la transaccion no valida (formato, doble cobro, fuera de regla).",
        "accion": "Verificar que no haya doble envio. Revisar reglas con la pasarela.",
        "responsable": "Pasarela",
    },
    {
        "match": "monto excede el máximo permitido por entidad",
        "es": "Monto excede el maximo permitido por la entidad (CAF)",
        "explica": "El banco tiene un cupo maximo por transaccion en la cuenta de cobros.",
        "accion": "Dividir el cobro en montos menores o gestionar aumento de CAF con la entidad.",
        "responsable": "Cliente",
    },
    {
        "match": "valor de cobrança foi zerado devido o desconto programado",
        "es": "Cobro en cero por descuento programado (BR)",
        "explica": "El cliente tiene un descuento programado que llevo el cobro a cero. No es un fallo real.",
        "accion": "Reclasificar como ajuste, no como fallo de debito.",
        "responsable": "Comercial",
    },
    {
        "match": "pick up card",
        "es": "Banco solicita retener la tarjeta",
        "explica": "El banco emisor pide retener el plastico (sospecha de fraude).",
        "accion": "Imposible recuperar. Contactar al cliente para nuevo medio de pago.",
        "responsable": "Cliente",
    },
    {
        "match": "expired card",
        "es": "Tarjeta vencida",
        "explica": "La fecha de vencimiento de la tarjeta ya paso.",
        "accion": "Pedir al cliente actualizar tarjeta. Considerar Account Updater para evitar caidas masivas.",
        "responsable": "Cliente",
    },
    {
        "match": "ocurrió un error inesperado",
        "es": "Error inesperado en la pasarela",
        "explica": "Error generico no clasificado.",
        "accion": "Reintentar. Si se repite, escalar a soporte de pasarela.",
        "responsable": "Pasarela",
    },
    {
        "match": "security violation",
        "es": "Violacion de seguridad detectada",
        "explica": "El banco detecto un patron sospechoso y rechazo el cobro.",
        "accion": "Pedir al cliente autorizar cobros recurrentes con su banco. Revisar antifraude de pasarela.",
        "responsable": "Cliente",
    },
    {
        "match": "undefined error",
        "es": "Error indefinido",
        "explica": "Respuesta no estandarizada del emisor.",
        "accion": "Reintentar. Loggear y monitorear con soporte de pasarela.",
        "responsable": "Pasarela",
    },
    {
        "match": "invalid account number",
        "es": "Numero de cuenta invalido",
        "explica": "El emisor no reconoce la cuenta asociada.",
        "accion": "Pedir al cliente actualizar el medio de pago.",
        "responsable": "Cliente",
    },
    {
        "match": "entidad no reportó la cuenta en archivo de saldos",
        "es": "Entidad no reporta la cuenta en archivo de saldos",
        "explica": "Problema de conciliacion; el banco no envio la cuenta del cliente al archivo de saldos.",
        "accion": "Contactar al banco emisor. Revisar conciliacion de archivo de saldos.",
        "responsable": "Banco",
    },
    {
        "match": "producto no está activo",
        "es": "Producto del cliente no esta activo en el banco",
        "explica": "La cuenta o tarjeta del socio aun no esta activa.",
        "accion": "Pedir al cliente activar el producto con su banco; reintentar luego.",
        "responsable": "Cliente",
    },
    {
        "match": "bloqueada por petición del cliente",
        "es": "Tarjeta bloqueada a peticion del cliente",
        "explica": "El cliente bloqueo la tarjeta voluntariamente.",
        "accion": "Contactar al cliente para nuevo medio de pago.",
        "responsable": "Cliente",
    },
    {
        "match": "valor não pode ser nulo",
        "es": "Valor no puede ser nulo",
        "explica": "Error de integracion: la pasarela recibio un valor nulo.",
        "accion": "Bug operativo: revisar pipeline de envio de cobros y validacion de monto.",
        "responsable": "Operacion",
    },
    {
        "match": "general system failure",
        "es": "Falla general del sistema",
        "explica": "Error sistemico de la pasarela o del banco.",
        "accion": "Reintentar. Escalar a soporte si es persistente.",
        "responsable": "Pasarela",
    },
    {
        "match": "unable to authorise - error",
        "es": "No fue posible autorizar (error generico)",
        "explica": "Error sin detalle del banco.",
        "accion": "Reintentar; si persiste, contactar al cliente.",
        "responsable": "Pasarela",
    },
    {
        "match": "authentication failure",
        "es": "Fallo de autenticacion (3DS)",
        "explica": "La autenticacion del titular (3D Secure) fallo.",
        "accion": "Cobros recurrentes no deberian requerir 3DS. Revisar reglas con la pasarela.",
        "responsable": "Pasarela",
    },
    {
        "match": "not permitted to cardholder",
        "es": "Operacion no permitida para el tarjetahabiente",
        "explica": "El banco no permite este tipo de transaccion para este cliente.",
        "accion": "Pedir al cliente autorizar cobros recurrentes con el banco.",
        "responsable": "Cliente",
    },
    {
        "match": "información errónea en el mensaje",
        "es": "Informacion erronea en el mensaje",
        "explica": "La pasarela envio campos incorrectos al banco.",
        "accion": "Bug operativo: revisar integracion con pasarela.",
        "responsable": "Operacion",
    },
    {
        "match": "invalid amount",
        "es": "Monto invalido",
        "explica": "Monto fuera del rango aceptado por el banco.",
        "accion": "Revisar configuracion de montos con pasarela y banco.",
        "responsable": "Operacion",
    },
    {
        "match": "function not supported",
        "es": "Funcion no soportada",
        "explica": "El banco no soporta el tipo de operacion solicitada.",
        "accion": "Pedir al cliente cambiar el medio de pago.",
        "responsable": "Cliente",
    },
    {
        "match": "límite de usos por período excedido",
        "es": "Limite de usos por periodo excedido",
        "explica": "El cliente alcanzo el numero maximo de usos de la tarjeta en el periodo.",
        "accion": "Reintentar al inicio del siguiente periodo. Pedir al cliente aumentar limite.",
        "responsable": "Cliente",
    },
    {
        "match": "issuing bank unavailable",
        "es": "Banco emisor no disponible",
        "explica": "El banco no respondio al cobro.",
        "accion": "Reintentar mas tarde.",
        "responsable": "Banco",
    },
    {
        "match": "lost or stolen card",
        "es": "Tarjeta perdida o robada",
        "explica": "El banco bloqueo la tarjeta por reporte de perdida o robo.",
        "accion": "Imposible recuperar. Contactar al cliente para nuevo medio de pago.",
        "responsable": "Cliente",
    },
    {
        "match": "insufficient funds",
        "es": "Fondos insuficientes",
        "explica": "Saldo insuficiente en la cuenta del cliente.",
        "accion": "Reintentar al cierre de mes o cuando se acredite la nomina.",
        "responsable": "Cliente",
    },
    {
        "match": "transacción duplicada",
        "es": "Transaccion duplicada",
        "explica": "La pasarela detecto un duplicado del cobro.",
        "accion": "Bug operativo: evitar doble envio del cobro.",
        "responsable": "Operacion",
    },
    {
        "match": "número de intentos del pin",
        "es": "PIN excedido (numero de intentos)",
        "explica": "El cliente excedio los intentos de PIN.",
        "accion": "Cobros automaticos no deberian pedir PIN; revisar config.",
        "responsable": "Pasarela",
    },
    {
        "match": "transação não enviada",
        "es": "Transaccion no enviada (BR)",
        "explica": "La transaccion no llego al banco por descuento programado u otro motivo.",
        "accion": "Reclasificar segun motivo. Si es por descuento, no es un fallo real.",
        "responsable": "Comercial",
    },
    {
        "match": "card has reached the credit limit",
        "es": "Tarjeta alcanzo el limite de credito",
        "explica": "Cupo de credito agotado.",
        "accion": "Pedir al cliente liberar cupo o cambiar de medio de pago.",
        "responsable": "Cliente",
    },
    {
        "match": "inactive card or card not authorized for card-not-present",
        "es": "Tarjeta inactiva o no autorizada para CNP",
        "explica": "El banco no autoriza tarjetas para transacciones sin presencia (CNP).",
        "accion": "Pedir al cliente activar para ecommerce/recurrente o cambiar de tarjeta.",
        "responsable": "Cliente",
    },
    {
        "match": "issuer off line",
        "es": "Emisor fuera de linea",
        "explica": "El banco emisor no esta respondiendo.",
        "accion": "Reintentar mas tarde.",
        "responsable": "Banco",
    },
    {
        "match": "too much usage",
        "es": "Uso excedido",
        "explica": "Limite de uso alcanzado por el cliente.",
        "accion": "Esperar reseteo del periodo. Pedir al cliente aumentar limite.",
        "responsable": "Cliente",
    },
    {
        "match": "mala selección de la cuenta o asociación tipo de tarjeta errado",
        "es": "Cuenta o tipo de tarjeta erroneo",
        "explica": "Datos de cuenta o tipo de tarjeta mal asociados.",
        "accion": "Operacion: revisar el alta del medio de pago.",
        "responsable": "Operacion",
    },
    {
        "match": "server timeout",
        "es": "Timeout del servidor",
        "explica": "La pasarela o el banco no respondio a tiempo.",
        "accion": "Reintentar inmediatamente.",
        "responsable": "Pasarela",
    },
    {
        "match": "unknown server error",
        "es": "Error desconocido del servidor",
        "explica": "Error no clasificado en la respuesta de la pasarela.",
        "accion": "Loggear, reintentar; escalar si es masivo.",
        "responsable": "Pasarela",
    },
    {
        "match": "monto inválido",
        "es": "Monto invalido",
        "explica": "Monto rechazado por la pasarela o el banco.",
        "accion": "Revisar configuracion de monto con pasarela.",
        "responsable": "Operacion",
    },
    {
        "match": "erro não especificado",
        "es": "Error no especificado (BR)",
        "explica": "Error generico devuelto por la pasarela brasilena.",
        "accion": "Reintentar. Si persiste, escalar a soporte.",
        "responsable": "Pasarela",
    },
    {
        "match": "processor declined",
        "es": "Procesador rechazo",
        "explica": "El procesador rechazo sin detalle.",
        "accion": "Reintentar; revisar reglas con pasarela.",
        "responsable": "Pasarela",
    },
    {
        "match": "ocorreu um erro inesperado",
        "es": "Ocurrio un error inesperado (BR)",
        "explica": "Error generico.",
        "accion": "Reintentar; escalar si persiste.",
        "responsable": "Pasarela",
    },
    {
        "match": "allowable pin tries exceeded",
        "es": "Intentos de PIN excedidos",
        "explica": "Cliente excedio intentos de PIN.",
        "accion": "Revisar configuracion: cobros recurrentes no deberian requerir PIN.",
        "responsable": "Pasarela",
    },
]


def lookup_motivo(motivo_text):
    """Devuelve dict con es/explica/accion/responsable o None."""
    if not isinstance(motivo_text, str):
        return None
    s = motivo_text.lower()
    for entry in MOTIVO_CATALOG:
        if entry["match"] in s:
            return entry
    return None


def translate_motivo(motivo_text):
    """Solo el texto en espanol (o el original si no hay match)."""
    e = lookup_motivo(motivo_text)
    if e:
        return e["es"]
    return motivo_text or "Sin motivo registrado"
