# -*- coding: utf-8 -*-
"""Cliente Gemini API para tareas complejas (router inteligente).
Usa google-generativeai (SDK oficial).
Modo gratuito: gemini-2.0-flash (1,500 req/dia sin costo).
"""
import os
import time
import logging
from typing import Optional

import google.generativeai as genai

log = logging.getLogger(__name__)


class GeminiClient:
    """Cliente para Gemini API con modo gratuito (flash) y premium (pro)."""

    MODELO_GRATIS = 'gemini-2.0-flash'
    MODELO_PREMIUM = 'gemini-2.5-pro'

    def __init__(self, api_key: Optional[str] = None, modelo: str = MODELO_GRATIS):
        self.api_key = api_key or os.getenv('GEMINI_API_KEY', '')
        if not self.api_key:
            log.warning('GEMINI_API_KEY no configurada. Gemini no estara disponible.')
        genai.configure(api_key=self.api_key)
        self.modelo = modelo
        self.generacion_config = {
            'temperature': 0.7,
            'top_p': 0.95,
            'top_k': 40,
            'max_output_tokens': 8192,
        }
        log.info(f'GeminiClient iniciado con modelo={modelo}')

    def cambiar_modelo(self, modelo: str):
        """Cambia entre gemini-2.0-flash (gratis) y gemini-2.5-pro (pago)."""
        if modelo in (self.MODELO_GRATIS, self.MODELO_PREMIUM):
            self.modelo = modelo
            log.info(f'Modelo Gemini cambiado a: {modelo}')

    def _modelo_disponible(self) -> bool:
        return bool(self.api_key)

    def generar(self, prompt: str, system_prompt: Optional[str] = None) -> str:
        """Genera respuesta con Gemini.
        
        Args:
            prompt: Consulta del usuario
            system_prompt: Instrucciones de sistema (Gemini lo recibe como parte del prompt)
        Returns:
            Texto de respuesta
        """
        if not self._modelo_disponible():
            return '[Gemini no disponible: API key no configurada]'

        try:
            model = genai.GenerativeModel(
                model_name=self.modelo,
                generation_config=self.generacion_config,
                system_instruction=system_prompt,
            )
            respuesta = model.generate_content(prompt)
            log.info(f'Gemini OK ({len(prompt)} chars prompt, {self.modelo})')
            return respuesta.text
        except Exception as e:
            log.error(f'Error en Gemini: {e}')
            return f'[Error Gemini: {str(e)}]'

    def generar_con_historial(self, mensajes: list, system_prompt: Optional[str] = None) -> str:
        """Genera respuesta con historial de conversacion.
        
        Args:
            mensajes: Lista de dicts [{'role': 'user'|'model', 'parts': ['texto']}]
            system_prompt: Instrucciones de sistema
        Returns:
            Texto de respuesta
        """
        if not self._modelo_disponible():
            return '[Gemini no disponible: API key no configurada]'

        try:
            model = genai.GenerativeModel(
                model_name=self.modelo,
                generation_config=self.generacion_config,
                system_instruction=system_prompt,
            )
            chat = model.start_chat(history=mensajes[:-1] if len(mensajes) > 1 else [])
            respuesta = chat.send_message(mensajes[-1]['parts'][0] if mensajes else '')
            log.info(f'Gemini chat OK ({len(mensajes)} mensajes, {self.modelo})')
            return respuesta.text
        except Exception as e:
            log.error(f'Error en Gemini chat: {e}')
            return f'[Error Gemini: {str(e)}]'


# Singleton para uso global
_cliente_gemini: Optional[GeminiClient] = None


def obtener_gemini() -> GeminiClient:
    """Obtiene o crea el singleton de GeminiClient."""
    global _cliente_gemini
    if _cliente_gemini is None:
        _cliente_gemini = GeminiClient()
    return _cliente_gemini
