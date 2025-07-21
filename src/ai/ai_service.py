# -*- coding: utf-8 -*-
"""
Enhanced AI Service Module for EchoScribe AI
Fully integrated from monolithic system with complete functionality.
Supports Groq Whisper transcription, AI text enhancement, and context detection.
"""

import logging
import time
import re
import json
import os
import threading
import tempfile
import wave
import audioop
from typing import Optional, Dict, Any, List, Tuple, Callable
from collections import Counter
import math

try:
    import groq
    GROQ_AVAILABLE = True
except ImportError:
    GROQ_AVAILABLE = False
    print("WARNING: Groq library not available. AI features will be disabled.")

try:
    import speech_recognition as sr
    SPEECH_RECOGNITION_AVAILABLE = True
except ImportError:
    SPEECH_RECOGNITION_AVAILABLE = False
    print("WARNING: speech_recognition library not available. Google Speech will be disabled.")

logger = logging.getLogger(__name__)

class AIService:
    """
    Enhanced AI service integrated from monolithic system.
    Supports transcription, enhancement, and context detection.
    """

    def __init__(self, api_key: Optional[str] = None, config: Optional[Dict] = None,
                 status_callback: Optional[Callable[[str], None]] = None):
        self.api_key = api_key
        self.config = config or {}
        self.status_callback = status_callback or (lambda x: None)
        self.client: Optional[groq.Groq] = None

        # Configuration from monolithic system
        self.use_economic_model = self.config.get('use_economic_model', False)
        self.chunk_size = self.config.get('chunk_size', 600)
        self.max_tokens = self.config.get('max_tokens', 4000)
        self.language = self.config.get('language', 'id-ID')
        self.engine = self.config.get('engine', 'Google')
        self.use_ai_enhancement = self.config.get('use_ai_enhancement', True)
        self.api_request_delay = self.config.get('api_request_delay', 10)

        # Rate limiting and threading
        self.api_semaphore = threading.Semaphore(3)
        self.last_api_call = 0

        # Speech recognition setup
        if SPEECH_RECOGNITION_AVAILABLE:
            self.recognizer = sr.Recognizer()

        # Content analysis patterns from monolithic
        self.content_patterns = {
            'meeting': [
                'agenda', 'meeting', 'rapat', 'diskusi', 'keputusan', 'action item',
                'follow up', 'deadline', 'assigned to', 'next meeting'
            ],
            'lecture': [
                'pembelajaran', 'materi', 'bab', 'topik', 'contoh', 'latihan',
                'tugas', 'ujian', 'quiz', 'presentation', 'slide'
            ],
            'interview': [
                'wawancara', 'interview', 'pertanyaan', 'jawaban', 'experience',
                'skills', 'qualification', 'background', 'motivation'
            ],
            'research': [
                'penelitian', 'research', 'data', 'analysis', 'hypothesis',
                'methodology', 'findings', 'conclusion', 'recommendation'
            ],
            'technical': [
                'sistem', 'teknologi', 'development', 'code', 'programming',
                'software', 'hardware', 'network', 'database', 'security'
            ]
        }

        # Initialize client if API key provided
        if api_key and GROQ_AVAILABLE:
            self._initialize_client()

    def _initialize_client(self) -> bool:
        """Initialize Groq client (from monolithic)."""
        try:
            if not self.api_key or not self.api_key.startswith("gsk_"):
                logger.error("Invalid Groq API key format")
                return False

            self.client = groq.Groq(api_key=self.api_key)
            logger.info("Groq client initialized successfully")
            self.status_callback("✅ AI service connected")
            return True

        except Exception as e:
            logger.error(f"Failed to initialize Groq client: {e}")
            self.client = None
            self.status_callback(f"❌ AI connection failed: {str(e)[:50]}...")
            return False

    def update_api_key(self, api_key: str) -> bool:
        """Update the API key and reinitialize the client."""
        self.api_key = api_key
        return self._initialize_client()

    def update_config(self, config: Dict[str, Any]) -> None:
        """Update service configuration."""
        self.config.update(config)
        self.use_economic_model = self.config.get('use_economic_model', False)
        self.chunk_size = self.config.get('chunk_size', 600)
        self.max_tokens = self.config.get('max_tokens', 4000)
        self.language = self.config.get('language', 'id-ID')
        self.engine = self.config.get('engine', 'Google')
        self.use_ai_enhancement = self.config.get('use_ai_enhancement', True)
        self.api_request_delay = self.config.get('api_request_delay', 10)

    def is_available(self) -> bool:
        """Check if AI service is available with enhanced validation."""
        if not GROQ_AVAILABLE:
            return False
        if self.client is None:
            return False
        # Test connection with a simple request
        try:
            # Quick validation call
            response = self.client.chat.completions.create(
                model="llama-3.1-8b-instant",
                messages=[{"role": "user", "content": "test"}],
                max_tokens=1
            )
            return True
        except Exception:
            return False

    def detect_audio_context(self, audio_file: str) -> Dict[str, Any]:
        """
        Detect audio characteristics and context for optimization.
        Integrated from monolithic system.
        """
        try:
            duration = self.get_audio_duration(audio_file)

            # Audio analysis with enhanced validation
            with wave.open(audio_file, 'rb') as wf:
                test_samplerate = wf.getframerate()
                if test_samplerate == 0:
                    logger.error(f"Invalid sample rate detected: {test_samplerate}")
                    return self._get_default_audio_context()

                # Get sample frames for analysis
                n_frames = min(wf.getnframes(), 1000000)
                frames = wf.readframes(n_frames)

                if not frames:
                    logger.warning("No audio frames found in file")
                    return self._get_default_audio_context()

                # Calculate RMS for average volume
                try:
                    rms = audioop.rms(frames, wf.getsampwidth())
                except audioop.error as audio_error:
                    logger.error(f"Audio processing error: {audio_error}")
                    return self._get_default_audio_context()

                # Calculate silence ratio
                silent_threshold = max(rms * 0.1, 100)
                silent_frames = 0
                frame_size = wf.getsampwidth() * wf.getnchannels()

                for i in range(0, len(frames), frame_size):
                    if i + frame_size <= len(frames):
                        chunk = frames[i:i + frame_size]
                        try:
                            chunk_rms = audioop.rms(chunk, wf.getsampwidth())
                            if chunk_rms < silent_threshold:
                                silent_frames += 1
                        except audioop.error:
                            continue

                silence_ratio = silent_frames / max(len(frames) / frame_size, 1)

            # Context detection
            context = {
                "duration": duration,
                "sample_rate": test_samplerate,
                "volume_level": "high" if rms > 10000 else "medium" if rms > 5000 else "low",
                "silence_ratio": silence_ratio,
                "content_type": self._detect_content_type(duration, silence_ratio),
                "audio_quality": "good" if rms > 1000 and test_samplerate >= 16000 else "poor"
            }

            logger.info(f"Audio context detected: {context}")
            return context

        except Exception as e:
            logger.error(f"Error in audio context detection: {e}")
            return self._get_default_audio_context()

    def _get_default_audio_context(self) -> Dict[str, Any]:
        """Return default audio context when detection fails."""
        return {
            "duration": 0,
            "sample_rate": 16000,
            "volume_level": "medium",
            "silence_ratio": 0,
            "content_type": "unknown",
            "audio_quality": "unknown"
        }

    def _detect_content_type(self, duration: float, silence_ratio: float) -> str:
        """Detect content type based on audio characteristics."""
        if duration > 3600:  # > 1 hour
            if silence_ratio > 0.3:
                return "lecture"
            else:
                return "meeting"
        elif duration > 1800:  # > 30 minutes
            if silence_ratio > 0.4:
                return "presentation"
            else:
                return "interview"
        else:
            if silence_ratio > 0.5:
                return "narrative"
            else:
                return "discussion"

    def get_audio_duration(self, audio_file: str) -> float:
        """Get audio duration from file."""
        try:
            with wave.open(audio_file, 'rb') as wf:
                frames = wf.getnframes()
                rate = wf.getframerate()
                return frames / rate if rate > 0 else 0.0
        except Exception as e:
            logger.error(f"Error getting audio duration: {e}")
            return 0.0

    def transcribe_audio_with_whisper(self, audio_file: str) -> str:
        """
        Transcribe audio using Groq Whisper.
        Integrated from monolithic system with full error handling.
        """
        try:
            if not self.is_available():
                return self._fallback_to_google_speech(audio_file)

            self.status_callback("🎤 Transcribing with Whisper AI...")

            # Rate limiting
            self._wait_for_rate_limit()

            # Get audio context for optimization
            context = self.detect_audio_context(audio_file)

            # Determine appropriate model based on context
            model = "whisper-large-v3" if not self.use_economic_model else "whisper-1"

            # Adjust parameters based on audio quality
            if context.get("audio_quality") == "poor":
                # Use more robust settings for poor quality audio
                prompt = "Ini adalah rekaman audio dengan kualitas suara yang mungkin tidak sempurna. "
            else:
                prompt = ""

            # Add language hint
            language_map = {
                "id-ID": "id",
                "en-US": "en",
                "ja-JP": "ja",
                "zh-CN": "zh"
            }
            language_code = language_map.get(self.language, "id")

            with open(audio_file, "rb") as file:
                transcription = self.client.audio.transcriptions.create(
                    file=(audio_file, file.read()),
                    model=model,
                    prompt=prompt,
                    response_format="text",
                    language=language_code,
                    temperature=0.0  # Deterministic output
                )

            self.status_callback("✅ Whisper transcription completed")
            return transcription.strip() if transcription else ""

        except Exception as e:
            logger.error(f"Whisper transcription error: {e}")
            self.status_callback(f"⚠️ Whisper failed, trying backup method...")
            return self._fallback_to_google_speech(audio_file)

    def _fallback_to_google_speech(self, audio_file: str) -> str:
        """Fallback to Google Speech Recognition."""
        try:
            if not SPEECH_RECOGNITION_AVAILABLE:
                self.status_callback("❌ No speech recognition available")
                return ""

            self.status_callback("🔄 Using Google Speech Recognition...")

            with sr.AudioFile(audio_file) as source:
                audio = self.recognizer.record(source)

            # Extract language code
            language_code = self.language if self.language != "id-ID" else "id"

            result = self.recognizer.recognize_google(audio, language=language_code)
            self.status_callback("✅ Google Speech transcription completed")
            return result

        except sr.UnknownValueError:
            self.status_callback("❌ Could not understand audio")
            return ""
        except sr.RequestError as e:
            self.status_callback(f"❌ Google Speech error: {str(e)[:50]}...")
            return ""
        except Exception as e:
            logger.error(f"Fallback transcription error: {e}")
            self.status_callback("❌ All transcription methods failed")
            return ""

    def enhance_text_with_ai(self, text: str, context: Optional[Dict[str, Any]] = None) -> str:
        """
        Enhanced AI text enhancement with advanced content analysis.
        Integrated from monolithic system with sophisticated optimization.
        """
        try:
            if not self.use_ai_enhancement or not self.is_available():
                return self._fallback_enhancement(text)

            if not text or len(text.strip()) < 10:
                return text

            self.status_callback("🤖 Analyzing content and enhancing with AI...")

            # Advanced content analysis
            content_stats = self._analyze_content_characteristics(text)

            # Select optimal model configuration
            model_config = self._select_optimal_model(content_stats)

            # Rate limiting
            self._wait_for_rate_limit()

            # Create sophisticated enhancement prompt
            enhancement_prompt = self._create_enhancement_prompt(text, content_stats.get('content_type', 'general'), context)

            self.status_callback(f"🔧 Using {model_config['name']} for {content_stats['content_type']} content...")

            completion = self.client.chat.completions.create(
                messages=[
                    {
                        "role": "system",
                        "content": f"Anda adalah {model_config['name']} yang ahli dalam menyusun catatan dan dokumentasi. "
                                  f"Tugas Anda adalah memperbaiki dan menyusun teks hasil transkripsi menjadi "
                                  f"dokumen yang terstruktur, profesional, dan mudah dibaca. "
                                  f"Konten yang akan diproses: {content_stats['content_type']}, "
                                  f"Kompleksitas: {content_stats.get('complexity', 'medium')}, "
                                  f"Jumlah kata: {content_stats.get('word_count', 0)}."
                    },
                    {
                        "role": "user",
                        "content": enhancement_prompt
                    }
                ],
                model=model_config["model_id"],
                temperature=model_config["temperature"],
                max_tokens=model_config["max_tokens"],
                top_p=model_config["top_p"]
            )

            enhanced_text = completion.choices[0].message.content

            # Post-process based on content characteristics
            final_text = self._post_process_enhanced_text(enhanced_text, content_stats)

            self.status_callback(f"✅ AI enhancement completed ({content_stats['content_type']})")

            return final_text.strip() if final_text else text

        except Exception as e:
            logger.error(f"AI enhancement error: {e}")
            self.status_callback(f"⚠️ AI enhancement failed, using fallback: {str(e)[:50]}...")
            return self._fallback_enhancement(text)

    def detect_language(self, text: str) -> str:
        """Detect the language of the given text."""
        try:
            # Simple language detection based on character patterns
            if not text:
                return self.language

            # Count common language indicators
            indonesian_words = ['yang', 'dan', 'dengan', 'untuk', 'dari', 'pada', 'adalah', 'akan', 'tidak', 'ini', 'itu']
            english_words = ['the', 'and', 'with', 'for', 'from', 'on', 'is', 'will', 'not', 'this', 'that']

            text_lower = text.lower()
            id_count = sum(1 for word in indonesian_words if word in text_lower)
            en_count = sum(1 for word in english_words if word in text_lower)

            if en_count > id_count:
                return "en-US"
            else:
                return "id-ID"

        except Exception as e:
            logger.error(f"Language detection error: {e}")
            return self.language

    def analyze_content_type(self, text: str) -> str:
        """Analyze text content to determine its type."""
        try:
            if not text:
                return "general"

            text_lower = text.lower()

            # Check for different content types based on keywords
            for content_type, keywords in self.content_patterns.items():
                matches = sum(1 for keyword in keywords if keyword in text_lower)
                if matches >= 2:  # At least 2 keywords match
                    return content_type

            return "general"

        except Exception as e:
            logger.error(f"Content type analysis error: {e}")
            return "general"

    def enhance_text(self, text: str, content_type: str = "general", language: str = "Indonesian") -> str:
        """Enhanced wrapper for text enhancement with content type and language support."""
        context = {
            "content_type": content_type,
            "language": language,
            "enhancement_level": "standard"
        }
        return self.enhance_text_with_ai(text, context)

    def transcribe_audio(self, audio_file: str, language: str = None, model: str = None) -> str:
        """Main transcription method with fallback support."""
        try:
            # Use Whisper if available, otherwise fallback to Google
            if self.is_available():
                return self.transcribe_audio_with_whisper(audio_file)
            else:
                return self._fallback_to_google_speech(audio_file)
        except Exception as e:
            logger.error(f"Transcription error: {e}")
            self.status_callback(f"Transcription failed: {str(e)[:50]}...")
            return ""

    def enhance_document_cohesion(self, text: str, content_type: str = "general", language: str = "Indonesian") -> str:
        """Enhance document cohesion for long, multi-chunk texts."""
        try:
            if not self.is_available():
                return text

            self.status_callback("🔗 Enhancing document cohesion...")

            # Rate limiting
            self._wait_for_rate_limit()

            # Create cohesion enhancement prompt
            cohesion_prompt = f"""
            Berikut adalah teks yang disusun dari beberapa bagian rekaman audio.
            Tugas Anda adalah memperbaiki alur dan kohesi keseluruhan dokumen dengan:

            1. Menghubungkan antar bagian dengan transisi yang natural
            2. Menghilangkan pengulangan yang tidak perlu
            3. Memastikan struktur yang logis dan mudah dibaca
            4. Memperbaiki format dan penomoran jika diperlukan

            Tipe konten: {content_type}
            Bahasa: {language}

            Teks asli:
            {text}

            Hasilkan dokumen yang telah diperbaiki kohesi dan strukturnya:
            """

            completion = self.client.chat.completions.create(
                messages=[
                    {
                        "role": "system",
                        "content": f"Anda adalah editor profesional yang ahli dalam menyusun dokumen yang kohesif dan terstruktur dalam bahasa {language}."
                    },
                    {
                        "role": "user",
                        "content": cohesion_prompt
                    }
                ],
                model="llama-3.1-70b-versatile",
                temperature=0.3,
                max_tokens=4000
            )

            enhanced_text = completion.choices[0].message.content
            self.status_callback("✅ Document cohesion enhanced")

            return enhanced_text.strip() if enhanced_text else text

        except Exception as e:
            logger.error(f"Document cohesion enhancement error: {e}")
            self.status_callback(f"⚠️ Cohesion enhancement failed: {str(e)[:50]}...")
            return text

    def _analyze_text_content(self, text: str) -> str:
        """Analyze text content to determine type."""
        text_lower = text.lower()

        # Count pattern matches
        pattern_scores = {}
        for content_type, patterns in self.content_patterns.items():
            score = sum(1 for pattern in patterns if pattern in text_lower)
            if score > 0:
                pattern_scores[content_type] = score

        # Return highest scoring type
        if pattern_scores:
            return max(pattern_scores, key=pattern_scores.get)
        else:
            return "general"

    def _create_enhancement_prompt(self, text: str, content_type: str, context: Optional[Dict] = None) -> str:
        """Create context-aware enhancement prompt."""
        base_prompt = f"""
Berikut adalah teks hasil transkripsi yang perlu diperbaiki dan disusun:

TEKS ASLI:
{text}

TUGAS ANDA:
1. Perbaiki ejaan, tata bahasa, dan tanda baca
2. Susun menjadi struktur yang logis dan mudah dibaca
3. Pertahankan semua informasi penting
4. Gunakan format yang sesuai untuk jenis konten: {content_type}
"""

        # Add specific instructions based on content type
        if content_type == "meeting":
            base_prompt += """
5. Buat struktur dengan bagian: Agenda, Pembahasan, Keputusan, Action Items
6. Gunakan bullet points untuk daftar
7. Highlight keputusan penting dan deadline
"""
        elif content_type == "lecture":
            base_prompt += """
5. Buat struktur dengan bagian: Topik Utama, Penjelasan, Contoh, Kesimpulan
6. Gunakan heading dan subheading yang jelas
7. Pisahkan konsep-konsep penting
"""
        elif content_type == "interview":
            base_prompt += """
5. Format sebagai Q&A yang jelas
6. Pisahkan pertanyaan dan jawaban
7. Highlight poin-poin penting dari jawaban
"""
        elif content_type == "technical":
            base_prompt += """
5. Gunakan struktur teknis yang logis
6. Pertahankan istilah teknis yang akurat
7. Buat daftar spesifikasi atau requirements jika ada
"""

        base_prompt += """

HASIL AKHIR:
Berikan teks yang sudah diperbaiki dan tersusun rapi, siap untuk dijadikan dokumen profesional.
"""

        return base_prompt

    def _analyze_content_characteristics(self, text: str) -> Dict[str, Any]:
        """
        Advanced content analysis with full characteristics detection.
        Integrated from monolithic system (lines 4380-4640).
        """
        try:
            words = text.split()
            lines = text.split('\n')

            # Basic statistics
            stats = {
                'word_count': len(words),
                'char_count': len(text),
                'line_count': len(lines),
                'paragraph_count': len([p for p in text.split('\n\n') if p.strip()]),
                'sentence_count': len([s for s in text.split('.') if s.strip()]),
                'avg_word_length': sum(len(w) for w in words) / max(len(words), 1),
                'avg_sentence_length': len(words) / max(len([s for s in text.split('.') if s.strip()]), 1)
            }

            # Content type detection with advanced patterns
            content_type = self._detect_content_type_advanced(text)
            stats['content_type'] = content_type

            # Language detection
            stats['language'] = self.detect_language(text)

            # Complexity analysis
            stats['complexity'] = self._calculate_complexity_score(text, words)

            # Structural analysis
            stats.update(self._analyze_structure(text))

            # Content-specific metrics
            if content_type == 'meeting':
                stats['meeting_metrics'] = self._analyze_meeting_content(text)
            elif content_type == 'lecture':
                stats['lecture_metrics'] = self._analyze_lecture_content(text)
            elif content_type == 'technical_report':
                stats['technical_metrics'] = self._analyze_technical_content(text)

            # Reading time estimation
            stats['reading_time_minutes'] = max(1, len(words) // 200)

            logger.info(f"Content analysis: {content_type}, {len(words)} words, complexity: {stats['complexity']}")
            return stats

        except Exception as e:
            logger.error(f"Error in content analysis: {e}")
            return {
                'content_type': 'general',
                'word_count': len(text.split()),
                'complexity': 'medium',
                'language': 'id'
            }

    def _detect_content_type_advanced(self, text: str) -> str:
        """
        Advanced content type detection with sophisticated patterns.
        Integrated from monolithic system.
        """
        text_lower = text.lower()

        # Advanced pattern matching
        patterns = {
            'meeting': [
                r'\b(rapat|meeting|agenda|keputusan|tindak lanjut|action item)\b',
                r'\b(pembahasan|diskusi|membahas|menyepakati)\b',
                r'\b(hadir|peserta|yang menghadiri|notulen)\b'
            ],
            'lecture': [
                r'\b(kuliah|lecture|pembelajaran|materi|bab|chapter)\b',
                r'\b(konsep|teori|definisi|penjelasan|understanding)\b',
                r'\b(mahasiswa|students|peserta didik|kelas)\b'
            ],
            'interview': [
                r'\b(wawancara|interview|pertanyaan|tanya|jawab)\b',
                r'\b(pengalaman|background|profile|cv)\b',
                r'\b(Q:|A:|pertanyaan|jawaban)\b'
            ],
            'technical_report': [
                r'\b(analisis|analysis|sistem|system|teknis|technical)\b',
                r'\b(spesifikasi|requirement|implementasi|configuration)\b',
                r'\b(database|server|API|framework|algorithm)\b'
            ],
            'presentation': [
                r'\b(presentasi|presentation|slide|demo|showcase)\b',
                r'\b(overview|ringkasan|summary|conclusion)\b',
                r'\b(target|goal|objective|hasil|result)\b'
            ]
        }

        scores = {}
        for content_type, pattern_list in patterns.items():
            score = 0
            for pattern in pattern_list:
                matches = len(re.findall(pattern, text_lower))
                score += matches
            scores[content_type] = score

        # Return highest scoring type
        if scores and max(scores.values()) > 0:
            return max(scores, key=scores.get)

        return "general"

    def _calculate_complexity_score(self, text: str, words: List[str]) -> str:
        """Calculate text complexity based on multiple factors."""
        try:
            # Technical terms count
            technical_terms = len(re.findall(r'\b[A-Z]{2,}\b', text))

            # Average sentence length
            sentences = [s for s in text.split('.') if s.strip()]
            avg_sentence_length = len(words) / max(len(sentences), 1)

            # Unique word ratio
            unique_words = len(set(word.lower() for word in words))
            unique_ratio = unique_words / max(len(words), 1)

            # Complexity scoring
            complexity_score = 0
            if avg_sentence_length > 20: complexity_score += 2
            elif avg_sentence_length > 15: complexity_score += 1

            if technical_terms > 10: complexity_score += 2
            elif technical_terms > 5: complexity_score += 1

            if unique_ratio > 0.7: complexity_score += 1

            if complexity_score >= 4:
                return "high"
            elif complexity_score >= 2:
                return "medium"
            else:
                return "low"

        except Exception:
            return "medium"

    def _analyze_structure(self, text: str) -> Dict[str, Any]:
        """Analyze document structure."""
        return {
            'has_headings': bool(re.search(r'^#+\s', text, re.MULTILINE)),
            'has_lists': bool(re.search(r'^\s*[-*+]\s', text, re.MULTILINE)),
            'has_numbered_lists': bool(re.search(r'^\s*\d+\.\s', text, re.MULTILINE)),
            'has_tables': bool(re.search(r'^\|.*\|$', text, re.MULTILINE)),
            'has_code': bool(re.search(r'```|`[^`]+`', text)),
            'has_quotes': bool(re.search(r'^>\s', text, re.MULTILINE)),
        }

    def _analyze_meeting_content(self, text: str) -> Dict[str, Any]:
        """Analyze meeting-specific content."""
        text_lower = text.lower()
        return {
            'has_agenda': any(word in text_lower for word in ['agenda', 'pembahasan', 'topik']),
            'has_action_items': any(word in text_lower for word in ['tindak lanjut', 'action', 'tugas']),
            'has_decisions': any(word in text_lower for word in ['keputusan', 'sepakat', 'decision']),
            'participant_mentions': len([w for w in text.split() if w.endswith(':') or '@' in w])
        }

    def _analyze_lecture_content(self, text: str) -> Dict[str, Any]:
        """Analyze lecture-specific content."""
        text_lower = text.lower()
        return {
            'has_concepts': any(word in text_lower for word in ['konsep', 'teori', 'concept', 'theory']),
            'has_examples': any(word in text_lower for word in ['contoh', 'misalnya', 'example', 'for instance']),
            'has_questions': text_lower.count('?') > 0,
            'educational_keywords': len([w for w in ['learn', 'study', 'understand', 'belajar', 'pahami'] if w in text_lower])
        }

    def _analyze_technical_content(self, text: str) -> Dict[str, Any]:
        """Analyze technical content."""
        text_lower = text.lower()
        technical_keywords = ['api', 'database', 'server', 'framework', 'algorithm', 'implementation']
        return {
            'technical_density': len([w for w in technical_keywords if w in text_lower]) / max(len(text.split()), 1),
            'has_code_blocks': '```' in text or '`' in text,
            'has_specifications': any(word in text_lower for word in ['spec', 'requirement', 'spesifikasi'])
        }

    def _select_optimal_model(self, content_stats: Dict[str, Any]) -> Dict[str, Any]:
        """
        Select optimal model and parameters based on content characteristics.
        Integrated from monolithic system (lines 4800+).
        """
        # Default configuration
        config = {
            "model_id": "deepseek-r1-distill-llama-70b",
            "temperature": 0.5,
            "max_tokens": 6000,
            "top_p": 0.95,
            "name": "Default AI"
        }

        content_type = content_stats.get('content_type', 'general')
        complexity = content_stats.get('complexity', 'medium')
        word_count = content_stats.get('word_count', 0)

        # Content type specific optimization
        if content_type == "technical_report":
            config.update({
                "temperature": 0.3,  # More deterministic for technical content
                "name": "Technical AI",
                "max_tokens": 8000 if word_count > 1000 else 6000
            })
        elif content_type == "meeting":
            config.update({
                "model_id": "deepseek-r1-distill-llama-70b",
                "temperature": 0.4,
                "name": "Meeting Notes AI"
            })
        elif content_type == "lecture":
            config.update({
                "temperature": 0.4,
                "name": "Educational AI"
            })
        elif content_type == "interview":
            config.update({
                "temperature": 0.5,
                "name": "Interview AI",
                "max_tokens": 7000
            })
        elif content_type == "narrative":
            config.update({
                "temperature": 0.6,  # More creative for narratives
                "name": "Narrative AI"
            })

        # Complexity adjustments
        if complexity == "high":
            config["max_tokens"] = min(config["max_tokens"] + 2000, 10000)
            config["temperature"] = max(config["temperature"] - 0.1, 0.2)
        elif complexity == "low":
            config["max_tokens"] = max(config["max_tokens"] - 1000, 3000)

        # Economic model consideration
        if self.use_economic_model and word_count < 1000:
            config["model_id"] = "llama3-8b-8192"
            config["name"] += " (Economic)"

        # Language-specific adjustments
        language = content_stats.get('language', 'id')
        if language == 'en':
            config["temperature"] += 0.05  # Slightly more flexible for English

        return config

    def _post_process_enhanced_text(self, enhanced_text: str, content_stats: Dict[str, Any]) -> str:
        """
        Post-process enhanced text based on content characteristics.
        Integrated from monolithic system.
        """
        try:
            # Basic cleanup
            text = enhanced_text.strip()

            # Remove common AI artifacts
            text = re.sub(r'^(Berikut|Here is|This is).*?:', '', text, flags=re.IGNORECASE)
            text = re.sub(r'^```.*?\n', '', text, flags=re.MULTILINE)
            text = re.sub(r'\n```.*?$', '', text, flags=re.MULTILINE)

            # Content-type specific post-processing
            content_type = content_stats.get('content_type', 'general')

            if content_type == 'meeting':
                # Ensure meeting structure
                if not re.search(r'(AGENDA|PEMBAHASAN|KESIMPULAN)', text, re.IGNORECASE):
                    text = self._add_meeting_structure(text)

            elif content_type == 'technical_report':
                # Ensure technical formatting
                text = self._enhance_technical_formatting(text)

            return text.strip()

        except Exception as e:
            logger.error(f"Error in post-processing: {e}")
            return enhanced_text

    def _add_meeting_structure(self, text: str) -> str:
        """Add meeting structure if missing."""
        if not re.search(r'^(AGENDA|PEMBAHASAN)', text, re.IGNORECASE | re.MULTILINE):
            return f"## PEMBAHASAN\n\n{text}"
        return text

    def _enhance_technical_formatting(self, text: str) -> str:
        """Enhance technical document formatting."""
        # Add code block formatting
        text = re.sub(r'(\w+\(\))', r'`\1`', text)
        # Add emphasis to technical terms
        text = re.sub(r'\b(API|URL|HTTP|JSON|XML|SQL)\b', r'**\1**', text)
        return text

    def _fallback_enhancement(self, text: str) -> str:
        """
        Fallback enhancement when AI is not available.
        Integrated from monolithic system.
        """
        try:
            # Basic text cleanup and formatting
            lines = text.split('\n')
            enhanced_lines = []

            for line in lines:
                line = line.strip()
                if not line:
                    enhanced_lines.append('')
                    continue

                # Basic sentence case correction
                if line and not line[0].isupper():
                    line = line[0].upper() + line[1:]

                # Add period if missing
                if line and line[-1] not in '.!?':
                    line += '.'

                enhanced_lines.append(line)

            # Join and clean up
            result = '\n'.join(enhanced_lines)
            result = re.sub(r'\n{3,}', '\n\n', result)  # Remove excessive newlines

            return result.strip()

        except Exception as e:
            logger.error(f"Error in fallback enhancement: {e}")
            return text

    def _wait_for_rate_limit(self):
        """Wait for rate limiting (from monolithic)."""
        try:
            current_time = time.time()
            time_since_last = current_time - self.last_api_call

            if time_since_last < self.api_request_delay:
                wait_time = self.api_request_delay - time_since_last
                logger.info(f"Rate limiting: waiting {wait_time:.1f} seconds")
                self.status_callback(f"⏳ Rate limiting: waiting {wait_time:.1f}s...")
                time.sleep(wait_time)

            self.last_api_call = time.time()

        except Exception as e:
            logger.error(f"Error in rate limiting: {e}")

    def transcribe_audio(self, audio_file: str) -> str:
        """
        Main transcription method with fallback support.
        """
        try:
            if not os.path.exists(audio_file):
                logger.error(f"Audio file not found: {audio_file}")
                return ""

            # Choose transcription engine
            if self.engine == "Whisper" and self.is_available():
                return self.transcribe_audio_with_whisper(audio_file)
            else:
                return self._fallback_to_google_speech(audio_file)

        except Exception as e:
            logger.error(f"Transcription error: {e}")
            self.status_callback(f"❌ Transcription failed: {str(e)[:50]}...")
            return ""

    def process_audio_chunks(self, audio_files: List[str]) -> str:
        """
        Process multiple audio chunks and combine results.
        From monolithic system's extended recording feature.
        """
        try:
            if not audio_files:
                return ""

            self.status_callback(f"🎤 Processing {len(audio_files)} audio chunks...")

            transcripts = []
            for i, audio_file in enumerate(audio_files):
                self.status_callback(f"🎤 Processing chunk {i+1}/{len(audio_files)}...")

                transcript = self.transcribe_audio(audio_file)
                if transcript.strip():
                    transcripts.append(transcript.strip())

                # Rate limiting between chunks
                if i < len(audio_files) - 1:
                    time.sleep(1)

            # Combine transcripts
            combined_text = " ".join(transcripts)

            if combined_text.strip():
                self.status_callback("✅ All chunks processed successfully")

                # Enhance combined text if enabled
                if self.use_ai_enhancement:
                    return self.enhance_text_with_ai(combined_text)
                else:
                    return combined_text
            else:
                self.status_callback("⚠️ No transcription could be generated")
                return ""

        except Exception as e:
            logger.error(f"Error processing audio chunks: {e}")
            self.status_callback(f"❌ Chunk processing failed: {str(e)[:50]}...")
            return ""

    def update_config(self, config: Dict[str, Any]):
        """Update configuration settings."""
        self.config.update(config)
        self.use_economic_model = config.get('use_economic_model', self.use_economic_model)
        self.chunk_size = config.get('chunk_size', self.chunk_size)
        self.max_tokens = config.get('max_tokens', self.max_tokens)
        self.language = config.get('language', self.language)
        self.engine = config.get('engine', self.engine)
        self.use_ai_enhancement = config.get('use_ai_enhancement', self.use_ai_enhancement)
        self.api_request_delay = config.get('api_request_delay', self.api_request_delay)

    def update_api_key(self, api_key: str) -> bool:
        """Update API key and reinitialize client."""
        try:
            if not api_key or not api_key.startswith("gsk_"):
                logger.error("Invalid API key format")
                self.status_callback("❌ Invalid API key format")
                return False

            old_api_key = self.api_key
            self.api_key = api_key

            if self._initialize_client():
                logger.info("API key updated successfully")
                self.status_callback("✅ AI service updated")
                return True
            else:
                # Rollback on failure
                self.api_key = old_api_key
                self._initialize_client()
                return False

        except Exception as e:
            logger.error(f"Error updating API key: {e}")
            self.status_callback("❌ Error updating API key")
            return False

    def detect_language(self, text: str) -> str:
        """
        Detect language of the given text.
        Simple implementation for testing purposes.
        """
        try:
            # Simple heuristic-based language detection
            text_lower = text.lower()

            # Indonesian patterns
            indonesian_words = ['dan', 'atau', 'yang', 'ini', 'itu', 'dengan', 'pada', 'untuk', 'dari', 'ke']
            indonesian_score = sum(1 for word in indonesian_words if word in text_lower)

            # English patterns
            english_words = ['the', 'and', 'or', 'this', 'that', 'with', 'for', 'from', 'to', 'in']
            english_score = sum(1 for word in english_words if word in text_lower)

            # Return detected language (short codes for testing)
            if indonesian_score > english_score:
                return 'id'
            elif english_score > 0:
                return 'en'
            else:
                return 'id'  # Default to Indonesian

        except Exception as e:
            logger.error(f"Error detecting language: {e}")
            return 'id'  # Default fallback

    def get_status(self) -> Dict[str, Any]:
        """Get current service status."""
        return {
            "available": self.is_available(),
            "groq_available": GROQ_AVAILABLE,
            "speech_recognition_available": SPEECH_RECOGNITION_AVAILABLE,
            "engine": self.engine,
            "language": self.language,
            "use_ai_enhancement": self.use_ai_enhancement,
            "use_economic_model": self.use_economic_model
        }
