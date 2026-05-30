import os
from openai import OpenAI
from typing import List, Dict, Any, Optional
import json
import httpx
from dotenv import load_dotenv
from .response_templates import ResponseTemplates, quick_response_fallback, detect_query_type, extract_manual_reference
from .response_cache import get_cached_response_fast, cache_response_fast, get_cache_stats

# Router Inteligente: Gemini como fallback para tareas complejas
try:
    from .gemini_client import obtener_gemini
    _GEMINI_DISPONIBLE = True
except ImportError:
    _GEMINI_DISPONIBLE = False

# Cargar variables de entorno
load_dotenv()

class DeepSeekClient:
    """Cliente para interactuar con DeepSeek API"""
    
    def __init__(self, api_key: str = None):
        self.api_key = api_key or os.getenv("DEEPSEEK_API_KEY")
        if not self.api_key:
            raise ValueError("DeepSeek API key no configurada")
        
        # Configurar cliente OpenAI compatible con DeepSeek
        # Crear cliente HTTP sin proxies
        http_client = httpx.Client(
            timeout=30.0,
            limits=httpx.Limits(max_keepalive_connections=5, max_connections=10)
        )
        
        self.client = OpenAI(
            api_key=self.api_key,
            base_url="https://api.deepseek.com",
            http_client=http_client
        )
    
    async def generate_response(
        self,
        prompt: str,
        context: List[Dict[str, str]] = None,
        temperature: float = 0.7,
        max_tokens: int = 1000
    ) -> str:
        """Generar respuesta usando DeepSeek"""
        
        messages = []
        
        # Añadir contexto si existe
        if context:
            messages.extend(context)
        
        # Añadir prompt actual
        messages.append({"role": "user", "content": prompt})
        
        try:
            response = self.client.chat.completions.create(
                model="deepseek-chat",
                messages=messages,
                temperature=temperature,
                max_tokens=max_tokens,
                stream=False
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            log_error = f"Error DeepSeek: {e}"
            print(log_error)
            # Fallback a Gemini si esta disponible
            if _GEMINI_DISPONIBLE:
                try:
                    gemini = obtener_gemini()
                    gemini_resp = gemini.generar(prompt)
                    if gemini_resp and not gemini_resp.startswith('[Error'):
                        return gemini_resp
                except Exception as ge:
                    print(f"Fallback Gemini tambien fallo: {ge}")
            return "Lo siento, estoy teniendo problemas para procesar tu solicitud. Por favor, intenta de nuevo."
    
    async def generate_response_with_knowledge(
        self,
        question: str,
        knowledge_context: str,
        conversation_history: List[Dict[str, str]] = None
    ) -> Dict[str, Any]:
        """Generar respuesta usando contexto de conocimiento con protocolo ITIL"""
        
        # 1. PRIMERO: Verificar caché para respuesta ultra-rápida
        cached_response = get_cached_response_fast(question, knowledge_context)
        if cached_response:
            print(f"⚡ Respuesta desde caché para: {question}")
            return cached_response
        
        # 2. Verificar si hay respuesta rápida disponible
        quick_response = quick_response_fallback(question)
        if quick_response:
            print(f"🚀 Usando respuesta rápida para: {question}")
            response = {
                "answer": quick_response,
                "needs_escalation": False,
                "confidence": 90
            }
            # Almacenar en caché para futuro
            cache_response_fast(question, response, knowledge_context)
            return response
        
        # 3. Detectar tipo de consulta para personalización
        query_type = detect_query_type(question)
        print(f"🔍 Tipo de consulta detectada: {query_type}")
        
        # 4. Extraer referencia de manual del contexto
        manual_ref = extract_manual_reference(knowledge_context)
        
        # 5. Debug: mostrar contexto recibido
        print(f"🔍 Pregunta: {question}")
        print(f"🔍 Longitud del contexto: {len(knowledge_context)} caracteres")
        print(f"🔍 Manual de referencia: {manual_ref}")
        
        # 6. Determinar si es primera interacción
        is_first_interaction = not conversation_history or len(conversation_history) == 0
        
        system_prompt = f"""Eres un **Agente de Mesa de Ayuda ITIL certificado** del sistema SIHCE-VIH. 
        Sigues protocolos profesionales de soporte técnico con respuestas estructuradas y amables.

        # CONTEXTO ACTUAL DE LA CONSULTA:
        • Tipo de consulta: {query_type.upper()}
        • Manual de referencia: {manual_ref}
        • Es primera interacción: {'Sí' if is_first_interaction else 'No'}

        # PROTOCOLO ITIL QUE DEBES SEGUIR:
        1. **SALUDO PROFESIONAL**: Comenzar con saludo apropiado
        2. **EMPATÍA**: Mostrar comprensión de la necesidad
        3. **PRECISIÓN**: Usar información exacta de manuales
        4. **ESTRUCTURA**: Organizar respuesta claramente
        5. **REFERENCIA**: Citar fuente específica
        6. **SEGUIMIENTO**: Ofrecer ayuda adicional

        # INFORMACIÓN DE LOS MANUALES DEL SISTEMA SIHCE-VIH:
        {knowledge_context}

        # FORMATO DE RESPUESTA ITIL OPTIMIZADO:
        [SALUDO EMPÁTICO] + [RECONOCIMIENTO CONSULTA] + [RESPUESTA ESTRUCTURADA] + [REFERENCIA] + [OFERTA VALOR AÑADIDO]

        # EJEMPLOS DE RESPUESTAS OPTIMIZADAS:
        
        EJEMPLO 1 - Procedimiento (con viñetas):
        "¡Buen día! Agente de Mesa de Ayuda SIHCE-VIH. 
        Entiendo que necesita conocer el procedimiento de registro PrEP. 
        
        📋 **Pasos según el manual:**
        1. Verificar ausencia de pruebas reactivas previas de VIH
        2. Acceder a Pantalla de Registro/Atención
        3. Seleccionar botón verde 'Registrar'
        4. Completar sección de Filiación
        
        📚 **Referencia:** Manual de Usuario v.1.4.8, sección 4.1
        
        💡 **¿Necesita que le detalle algún paso específico o tiene alguna duda durante el proceso?**"

        EJEMPLO 2 - Información (con formato claro):
        "¡Hola! Mesa de Ayuda SIHCE-VIH a su servicio. 
        Sobre su consulta acerca del sistema:
        
        📖 **Información del sistema:**
        El SIHCE-VIH es la plataforma integral para el registro y seguimiento de pacientes con VIH, incluyendo módulos de Prevención Combinada y reportes estadísticos.
        
        📋 **Fuente:** Manual de Usuario v.1.1.2, introducción
        
        🔍 **¿Le interesa conocer algún módulo o funcionalidad en particular?**"

        EJEMPLO 3 - Sin información (con alternativas):
        "Buen día, soy su agente de soporte. 
        He revisado la documentación disponible y no encuentro información específica sobre ese tema.
        
        🔄 **Alternativas recomendadas:**
        1. Contactar al administrador del sistema
        2. Consultar documentación técnica interna
        3. Escalar a soporte especializado
        
        💡 **¿Puedo ayudarle con algún otro procedimiento del sistema mientras tanto?**"

        # REGLAS DE CALIDAD ITIL:
        1. Tono profesional pero amable
        2. Estructura clara con encabezados cuando corresponda
        3. Información verificada en manuales
        4. Referencias específicas a documentos
        5. Oferta de valor añadido al final
        6. Uso de emojis relevantes para mejor legibilidad 📋🔍💡
        7. Evitar jerga técnica excesiva

        HISTORIAL DE CONVERSACIÓN (si aplica):
        {json.dumps(conversation_history or [], ensure_ascii=False)}
        """
        
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": question}
        ]
        
        try:
            response = self.client.chat.completions.create(
                model="deepseek-chat",
                messages=messages,
                temperature=0.3,  # Baja temperatura para respuestas más precisas
                max_tokens=1500
            )
            
            answer = response.choices[0].message.content
            
            # Analizar si necesita escalamiento
            needs_escalation = self._check_escalation_needed(question, answer)
            
            # Crear respuesta completa
            final_response = {
                "answer": answer,
                "needs_escalation": needs_escalation,
                "confidence": self._estimate_confidence(question, answer, knowledge_context)
            }
            
            # Almacenar en caché para respuestas futuras (solo si es buena respuesta)
            if final_response["confidence"] > 70 and not needs_escalation:
                cache_response_fast(question, final_response, knowledge_context)
                print(f"💾 Respuesta almacenada en caché (confianza: {final_response['confidence']})")
            
            return final_response
            
        except Exception as e:
            print(f"Error en respuesta con conocimiento: {e}")
            return {
                "answer": "Lo siento, estoy teniendo problemas para acceder a la información. Por favor, intenta de nuevo o contacta con un asistente humano.",
                "needs_escalation": True,
                "confidence": 0
            }
    
    def _check_escalation_needed(self, question: str, answer: str) -> bool:
        """Determinar si la conversación necesita escalamiento a humano"""
        
        escalation_triggers = [
            "no tengo información",
            "no sé",
            "no puedo ayudarte con eso",
            "contacta con un asistente",
            "habla con un humano",
            "escala la conversación",
            "problema complejo",
            "error del sistema",
            "queja",
            "reclamo"
        ]
        
        question_lower = question.lower()
        answer_lower = answer.lower()
        
        # Verificar triggers en la respuesta
        for trigger in escalation_triggers:
            if trigger in answer_lower:
                return True
        
        # Verificar si la pregunta parece compleja (muchas palabras)
        if len(question.split()) > 50:
            return True
        
        # Verificar palabras clave de emergencia en la pregunta
        emergency_keywords = ["urgente", "emergencia", "crítico", "no funciona", "error grave"]
        for keyword in emergency_keywords:
            if keyword in question_lower:
                return True
        
        return False
    
    def _estimate_confidence(
        self, 
        question: str, 
        answer: str, 
        knowledge_context: str
    ) -> int:
        """Estimar confianza de la respuesta (0-100)"""
        
        confidence = 70  # Base
        
        # Ajustar basado en longitud de respuesta
        if len(answer.split()) < 10:
            confidence -= 20  # Respuesta muy corta
        
        # Ajustar si menciona "no sé" o similar
        if any(phrase in answer.lower() for phrase in ["no sé", "no tengo", "no puedo"]):
            confidence -= 30
        
        # Ajustar si la respuesta cita el contexto
        if knowledge_context and any(
            word in answer for word in knowledge_context.split()[:20]
        ):
            confidence += 15
        
        # Asegurar límites
        return max(0, min(100, confidence))
    
    async def analyze_conversation_pattern(
        self,
        conversation_messages: List[Dict[str, str]]
    ) -> Dict[str, Any]:
        """Analizar patrones en la conversación para aprendizaje"""
        
        if not conversation_messages:
            return {"patterns": [], "suggestions": []}
        
        analysis_prompt = f"""Analiza esta conversación y extrae:
        1. Patrones de preguntas comunes
        2. Respuestas efectivas
        3. Puntos donde se necesitó escalar a humano
        4. Sugerencias para mejorar respuestas futuras

        CONVERSACIÓN:
        {json.dumps(conversation_messages, ensure_ascii=False)}

        Devuelve JSON con: patterns, suggestions, escalation_triggers
        """
        
        try:
            response = self.client.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": analysis_prompt}],
                temperature=0.1,
                max_tokens=2000,
                response_format={"type": "json_object"}
            )
            
            return json.loads(response.choices[0].message.content)
            
        except Exception as e:
            print(f"Error analizando patrones: {e}")
            return {"patterns": [], "suggestions": [], "escalation_triggers": []}