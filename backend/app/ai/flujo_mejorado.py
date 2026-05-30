"""
Flujo MEJORADO - Conocimiento estructurado + Tickets automáticos después de 2 negativas
"""

from typing import Dict, Any
from datetime import datetime
import re
from sqlalchemy.orm import Session
from app.database.models import UserProfile
from app.knowledge.conocimiento_estructurado import obtener_respuesta_para_consulta

# Router Inteligente: DeepSeek para rapido, Gemini para complejo
try:
    from .gemini_client import obtener_gemini
    _GEMINI_DISPONIBLE = True
except ImportError:
    _GEMINI_DISPONIBLE = False


class FlujoMejorado:
    """Flujo mejorado con conocimiento estructurado y tickets automáticos"""
    
    # Palabras clave para tareas complejas que requieren Gemini
    _COMPLEJAS = re.compile(r'analisis?|reporte|resumen?|compara|\bgrafic|\bestadistic|'
                            r'\btendencia|\bevolucion|\bpromedio|\bcalidad|\bbrecha|'
                            r'\bvinculacion|escenario|\bsimulac|\bproyect|'
                            r'\bhistorico|\bmensual|\btrimestral|\banual|'
                            r'\bexcel|\barchivo|\badjunt|\bdescarg|'
                            r'\bBD\b|\bdb\b|consulta sql|\bquery|\btabla', re.IGNORECASE)

    def __init__(self, db_session: Session = None):
        self.db = db_session
        self.estados_temporales = {}
        self.opciones_menu = {
            "1": "indicador tptb",
            "2": "establecimientos por región", 
            "3": "regularización de derivaciones",
            "4": "actualizar datos de paciente",
            "5": "problemas con aplicativo tar"
        }
    
    def set_db(self, db_session: Session):
        """Establecer sesión de base de datos"""
        self.db = db_session
    
    def procesar_mensaje(self, whatsapp_id: str, phone_number: str, 
                        user_input: str) -> Dict[str, Any]:
        """Procesar mensaje - Versión MEJORADA"""
        
        print(f"🔄 FlujoMejorado para {whatsapp_id}")
        
        if not self.db:
            return self._error("No hay conexión a base de datos")
        
        # 1. Buscar perfil en PostgreSQL
        perfil = self._obtener_o_crear_perfil(whatsapp_id, phone_number)
        
        # 2. Determinar estado
        estado = self._determinar_estado(whatsapp_id, perfil)
        print(f"   Estado: {estado}, Perfil completo: {perfil.profile_complete}")
        
        # 3. Procesar
        if estado == "INICIO":
            return self._fase_inicio(whatsapp_id, perfil, user_input)
        elif estado == "ESPERANDO_DATOS":
            return self._fase_datos(whatsapp_id, perfil, user_input)
        elif estado == "LISTO_PARA_CONSULTA":
            return self._fase_consulta(whatsapp_id, perfil, user_input)
        elif estado == "MOSTRO_MENU":
            return self._fase_menu(whatsapp_id, perfil, user_input)
        elif estado == "SATISFACCION":
            return self._fase_satisfaccion(whatsapp_id, perfil, user_input)
        else:
            return self._reiniciar(whatsapp_id, perfil)
    
    def _obtener_o_crear_perfil(self, whatsapp_id: str, phone_number: str) -> UserProfile:
        """Obtener o crear perfil en PostgreSQL"""
        perfil = self.db.query(UserProfile).filter(
            UserProfile.whatsapp_id == whatsapp_id
        ).first()
        
        if not perfil:
            perfil = UserProfile(
                whatsapp_id=whatsapp_id,
                phone_number=phone_number,
                user_name="",
                diris="",
                establishment="",
                profile_complete=0,
                last_interaction=datetime.now()
            )
            self.db.add(perfil)
            self.db.commit()
            self.db.refresh(perfil)
            print(f"   ✅ Nuevo perfil creado en DB")
        
        return perfil
    
    def _determinar_estado(self, whatsapp_id: str, perfil: UserProfile) -> str:
        """Determinar estado actual"""
        # Si ya tenemos estado en memoria para esta sesión
        if whatsapp_id in self.estados_temporales:
            return self.estados_temporales[whatsapp_id]
        
        # Si el perfil ya está completo en DB
        if perfil.profile_complete >= 3 and perfil.user_name and perfil.diris and perfil.establishment:
            return "LISTO_PARA_CONSULTA"
        else:
            return "INICIO"
    
    def _actualizar_estado(self, whatsapp_id: str, estado: str):
        """Actualizar estado en memoria"""
        self.estados_temporales[whatsapp_id] = estado
    
    def _guardar_datos_perfil(self, perfil: UserProfile, nombre: str, diris: str, establecimiento: str):
        """Guardar datos en perfil de PostgreSQL"""
        perfil.user_name = nombre
        perfil.diris = diris
        perfil.establishment = establecimiento
        perfil.profile_complete = 3
        perfil.last_interaction = datetime.now()
        self.db.commit()
        print(f"   💾 Perfil actualizado en DB: {nombre}")
    
    def _fase_inicio(self, whatsapp_id: str, perfil: UserProfile, user_input: str) -> Dict[str, Any]:
        """Fase inicial - Verificar si ya tiene datos"""
        
        # Si ya tiene datos en DB, saltar directamente a consulta
        if perfil.profile_complete >= 3 and perfil.user_name:
            self._actualizar_estado(whatsapp_id, "LISTO_PARA_CONSULTA")
            return {
                "answer": f"""¡Hola de nuevo, {perfil.user_name}! 👋

📍 DIRIS: {perfil.diris}
🏥 Establecimiento: {perfil.establishment}

¿En qué puedo ayudarte hoy?""",
                "needs_escalation": False,
                "confidence": 100,
                "modified": True,
                "fase": "CONSULTA"
            }
        
        # Si no tiene datos, pedirlos
        self._actualizar_estado(whatsapp_id, "ESPERANDO_DATOS")
        return {
            "answer": """📋 *Para comenzar, necesito registrar sus datos básicos:*
1. Nombre completo
2. DIRIS/Región
3. Establecimiento de salud

*Responda en el formato:*
- [Nombre]
- [DIRIS]
- [Establecimiento]*""",
            "needs_escalation": False,
            "confidence": 100,
            "modified": True,
            "fase": "DATOS"
        }
    
    def _fase_datos(self, whatsapp_id: str, perfil: UserProfile, user_input: str) -> Dict[str, Any]:
        """Procesar datos del usuario"""
        
        lines = [line.strip() for line in user_input.split('\n') if line.strip()]
        
        if len(lines) >= 3:
            nombre = lines[0]
            diris = lines[1]
            establecimiento = lines[2]
            
            # Guardar en PostgreSQL
            self._guardar_datos_perfil(perfil, nombre, diris, establecimiento)
            
            self._actualizar_estado(whatsapp_id, "LISTO_PARA_CONSULTA")
            
            return {
                "answer": f"""✅ *Registro completado, {nombre}.*
📍 DIRIS: {diris}
🏥 Establecimiento: {establecimiento}

¿Cuál es su consulta?""",
                "needs_escalation": False,
                "confidence": 100,
                "modified": True,
                "fase": "CONSULTA"
            }
        else:
            return {
                "answer": """❌ *Formato incorrecto.*

Por favor, responda exactamente:
- [Nombre completo]
- [DIRIS/Región]
- [Establecimiento de salud]

*Ejemplo:*
- Juan Pérez
- Lima Centro
- Hospital Regional""",
                "needs_escalation": False,
                "confidence": 100,
                "modified": True,
                "fase": "DATOS"
            }
    
    # ── Router Inteligente ──────────────────────────────────────────
    def _elegir_modelo(self, consulta: str) -> str:
        """Router: DeepSeek para rapido, Gemini para tareas complejas.
        Retorna 'deepseek' o 'gemini'."""
        if _GEMINI_DISPONIBLE and self._COMPLEJAS.search(consulta):
            return 'gemini'
        return 'deepseek'

    def _generar_con_router(self, consulta: str, contexto: str = '') -> str:
        """Genera respuesta usando el router inteligente.
        DeepSeek = rapido, Gemini = analisis complejo, fallback mutuo."""
        modelo = self._elegir_modelo(consulta)
        prompt = f"{contexto}\n\nConsulta: {consulta}" if contexto else consulta

        if modelo == 'gemini':
            try:
                gemini = obtener_gemini()
                system = 'Eres un asistente de ayuda para personal de salud del programa VIH/SIDA. Responde en espanol de forma clara y concisa.'
                resp = gemini.generar(prompt, system_prompt=system)
                if resp and not resp.startswith('[Error'):
                    return resp
            except Exception:
                pass
            # Fallback a DeepSeek si Gemini falla

        # Usar DeepSeek via OpenAI
        from .deepseek_client import DeepSeekClient
        try:
            ds = DeepSeekClient()
            import asyncio
            resp = asyncio.run(ds.generate_response(prompt, temperature=0.7, max_tokens=2000))
            if resp and 'problemas para procesar' not in resp:
                return resp
        except Exception:
            pass

        # Fallback final a Gemini si DeepSeek fallo
        if modelo == 'deepseek' and _GEMINI_DISPONIBLE:
            try:
                gemini = obtener_gemini()
                resp = gemini.generar(prompt)
                if resp and not resp.startswith('[Error'):
                    return resp
            except Exception:
                pass

        return obtener_respuesta_para_consulta(consulta)

    def _fase_consulta(self, whatsapp_id: str, perfil: UserProfile, user_input: str) -> Dict[str, Any]:
        """Procesar consulta - CON CONOCIMIENTO ESTRUCTURADO"""
        
        # Verificar si es saludo
        user_input_lower = user_input.lower().strip()
        saludos = ["hola", "buenos días", "buenas tardes", "buenas noches", "hi", "hello"]
        
        if any(saludo in user_input_lower for saludo in saludos):
            # Es saludo
            return {
                "answer": f"""¡Hola {perfil.user_name}! 👋

📍 DIRIS: {perfil.diris}
🏥 Establecimiento: {perfil.establishment}

¿En qué puedo ayudarte hoy?""",
                "needs_escalation": False,
                "confidence": 100,
                "modified": True,
                "fase": "CONSULTA"
            }
        
        # Guardar consulta actual
        self.estados_temporales[f"{whatsapp_id}_consulta"] = user_input
        self.estados_temporales[f"{whatsapp_id}_intentos"] = 0  # Inicializar contador de intentos
        
        # Verificar si es consulta específica
        if self._es_consulta_especifica(user_input):
            # Obtener respuesta del conocimiento estructurado
            respuesta = obtener_respuesta_para_consulta(user_input)
            
            self._actualizar_estado(whatsapp_id, "SATISFACCION")
            
            return {
                "answer": f"""{respuesta}

*¿La respuesta fue satisfactoria?*
1. ✅ Sí, quedó resuelta mi consulta
2. ❌ No, necesito más detalles o no era lo que buscaba""",
                "needs_escalation": False,
                "confidence": 85,
                "modified": True,
                "fase": "SATISFACCION"
            }
        else:
            # Intentar con router inteligente antes del menú
            try:
                contexto = f"Usuario: {perfil.user_name}, DIRIS: {perfil.diris}, Establecimiento: {perfil.establishment}"
                respuesta_router = self._generar_con_router(user_input, contexto)
                if respuesta_router and 'problemas para procesar' not in respuesta_router.lower():
                    self._actualizar_estado(whatsapp_id, "SATISFACCION")
                    return {
                        "answer": f"""{respuesta_router}

*¿La respuesta fue satisfactoria?*
1. ✅ Sí, quedó resuelta mi consulta
2. ❌ No, necesito más detalles o no era lo que buscaba""",
                        "needs_escalation": False,
                        "confidence": 80,
                        "modified": True,
                        "fase": "SATISFACCION"
                    }
            except Exception:
                pass

            # Fallback: mostrar menú
            self._actualizar_estado(whatsapp_id, "MOSTRO_MENU")
            return {
                "answer": f"""No estoy seguro de entender su consulta: "{user_input}"

¿A cuál de estos temas se refiere?

1. Información sobre indicador TPTB
2. Listado de establecimientos por región
3. Regularización de derivaciones
4. Actualizar datos de paciente
5. Problemas con el aplicativo TAR

*Responda con el número de la opción (1-5).*""",
                "needs_escalation": False,
                "confidence": 60,
                "modified": True,
                "fase": "CONSULTA_MENU"
            }
    
    def _fase_menu(self, whatsapp_id: str, perfil: UserProfile, user_input: str) -> Dict[str, Any]:
        """Procesar selección del menú"""
        
        user_input_lower = user_input.lower().strip()
        
        if user_input_lower in self.opciones_menu:
            consulta = self.opciones_menu[user_input_lower]
            self.estados_temporales[f"{whatsapp_id}_consulta"] = consulta
            self.estados_temporales[f"{whatsapp_id}_intentos"] = 0  # Inicializar contador
            
            # Obtener respuesta del conocimiento estructurado
            respuesta = obtener_respuesta_para_consulta(consulta)
            
            self._actualizar_estado(whatsapp_id, "SATISFACCION")
            
            return {
                "answer": f"""{respuesta}

*¿La respuesta fue satisfactoria?*
1. ✅ Sí, quedó resuelta mi consulta
2. ❌ No, necesito más detalles o no era lo que buscaba""",
                "needs_escalation": False,
                "confidence": 85,
                "modified": True,
                "fase": "SATISFACCION"
            }
        else:
            # No es número, tratar como nueva consulta
            self._actualizar_estado(whatsapp_id, "LISTO_PARA_CONSULTA")
            return self._fase_consulta(whatsapp_id, perfil, user_input)
    
    def _fase_satisfaccion(self, whatsapp_id: str, perfil: UserProfile, user_input: str) -> Dict[str, Any]:
        """Evaluar satisfacción - CON TICKET AUTOMÁTICO DESPUÉS DE 2 NEGATIVAS"""
        
        user_input_lower = user_input.lower().strip()
        
        # Inicializar contador si no existe
        if f"{whatsapp_id}_intentos" not in self.estados_temporales:
            self.estados_temporales[f"{whatsapp_id}_intentos"] = 0
        
        if user_input_lower in ["1", "sí", "si", "yes"]:
            # ✅ USUARIO SATISFECHO - Crear ticket normal
            consulta = self.estados_temporales.get(f"{whatsapp_id}_consulta", "No especificada")
            ticket_numero = f"TICKET-{datetime.now().year}-{perfil.id:05d}"
            
            self._actualizar_estado(whatsapp_id, "LISTO_PARA_CONSULTA")
            
            return {
                "answer": f"""✅ *Ticket registrado exitosamente.*

📋 **Número de ticket:** {ticket_numero}
👤 **Solicitante:** {perfil.user_name}
📍 **DIRIS:** {perfil.diris}
🏥 **Establecimiento:** {perfil.establishment}
📝 **Consulta:** {consulta}
📊 **Estado:** Resuelto por el sistema

*¿Necesita ayuda con algo más?*""",
                "needs_escalation": False,
                "confidence": 100,
                "modified": True,
                "fase": "CONSULTA"
            }
        elif user_input_lower in ["2", "no"]:
            # ❌ USUARIO NO SATISFECHO - Incrementar intentos
            self.estados_temporales[f"{whatsapp_id}_intentos"] += 1
            intentos = self.estados_temporales[f"{whatsapp_id}_intentos"]
            consulta = self.estados_temporales.get(f"{whatsapp_id}_consulta", "")
            
            print(f"   ⚠️ Intento #{intentos} de insatisfacción para {whatsapp_id}")
            
            if intentos >= 2:
                # 🚨 SEGUNDA NEGATIVA - Crear ticket con ESCALACIÓN AUTOMÁTICA
                ticket_numero = f"TICKET-{datetime.now().year}-{perfil.id:05d}-ESC"
                
                self._actualizar_estado(whatsapp_id, "LISTO_PARA_CONSULTA")
                
                return {
                    "answer": f"""🔧 *Ticket con ESCALACIÓN AUTOMÁTICA a analista*

📋 **Número de ticket:** {ticket_numero}
👤 **Solicitante:** {perfil.user_name}
📍 **DIRIS:** {perfil.diris}
🏥 **Establecimiento:** {perfil.establishment}
📝 **Consulta:** {consulta}
⚠️ **Motivo:** 2 respuestas insatisfactorias
👨‍💼 **Asignado a:** Analista según horario
📧 **Notificación:** Enviada al equipo

*Su consulta ha sido ESCALADA a un analista humano.*""",
                    "needs_escalation": True,
                    "confidence": 100,
                    "modified": True,
                    "fase": "CONSULTA"
                }
            else:
                # Primera negativa - Profundizar
                return {
                    "answer": f"""Voy a profundizar en su consulta sobre: {consulta}

*Información adicional:*
- Revise la documentación oficial
- Contacte al soporte técnico especializado
- Consulte con el jefe inmediato

¿Esta información adicional le resulta útil?
(En caso negativo, se generará ticket con escalación automática)""",
                    "needs_escalation": False,
                    "confidence": 75,
                    "modified": True,
                    "fase": "SATISFACCION"
                }
        else:
            return {
                "answer": """No entendí su respuesta.

*Por favor, responda con:*
1. ✅ Sí, quedó resuelta mi consulta
2. ❌ No, necesito más detalles

*Nota:* Después de 2 respuestas negativas, se generará ticket con escalación automática.""",
                "needs_escalation": False,
                "confidence": 100,
                "modified": True,
                "fase": "SATISFACCION"
            }
    
    def _reiniciar(self, whatsapp_id: str, perfil: UserProfile) -> Dict[str, Any]:
        """Reiniciar flujo"""
        self._actualizar_estado(whatsapp_id, "INICIO")
        return self._fase_inicio(whatsapp_id, perfil, "")
    
    def _error(self, mensaje: str) -> Dict[str, Any]:
        """Mensaje de error"""
        return {
            "answer": f"⚠️ *Error del sistema:* {mensaje}",
            "needs_escalation": False,
            "confidence": 100,
            "modified": True,
            "fase": "ERROR"
        }
    
    def _es_consulta_especifica(self, user_input: str) -> bool:
        """Determinar si es consulta específica"""
        consultas_especificas = [
            "exclusiones indicador 10",
            "gestantes", 
            "tptb",
            "indicador 10",
            "renipres",
            "establecimientos",
            "derivaciones",
            "actualizar datos",
            "aplicativo tar",
            "problemas técnicos",
            "sistema",
            "manual",
            "procedimiento",
            "indicador",
            "tratamiento",
            "preventivo"
        ]
        
        user_input_lower = user_input.lower()
        for consulta in consultas_especificas:
            if consulta in user_input_lower:
                return True
        return False
