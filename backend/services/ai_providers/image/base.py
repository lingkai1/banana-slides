"""
Abstract base class for image generation providers
"""
from abc import ABC, abstractmethod
from typing import Optional, List
from PIL import Image


class ImageProvider(ABC):
    """Abstract base class for image generation"""
    
    @abstractmethod
    def generate_image(
        self,
        prompt: str,
        ref_images: Optional[List[Image.Image]] = None,
        aspect_ratio: str = "16:9",
        resolution: str = "2K",
        project_id: str = None,
        page_id: str = None
    ) -> Optional[Image.Image]:
        """
        Generate image from prompt
        
        Args:
            prompt: The image generation prompt
            ref_images: Optional list of reference images (PIL Image objects)
            aspect_ratio: Image aspect ratio (e.g., "16:9", "1:1", "4:3")
            resolution: Image resolution ("1K", "2K", "4K") - note: OpenAI format only supports 1K
            project_id: Optional project ID (required for PPT Agent)
            page_id: Optional page ID (required for PPT Agent)
            
        Returns:
            Generated PIL Image object, or None if failed
        """
        pass
