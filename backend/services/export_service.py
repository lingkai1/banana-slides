"""
Export Service - handles PPTX and PDF export
Based on demo.py create_pptx_from_images()
"""
import os
import logging
from pathlib import Path
from typing import List
from pptx import Presentation
from pptx.util import Inches
from PIL import Image
import io
from pptx.enum.shapes import MSO_SHAPE_TYPE

logger = logging.getLogger(__name__)


class ExportService:
    """Service for exporting presentations"""
    
    @staticmethod
    def create_pptx_from_images(image_paths: List[str], output_file: str = None) -> bytes:
        """
        Create PPTX file from image paths
        Based on demo.py create_pptx_from_images()
        
        Args:
            image_paths: List of absolute paths to images
            output_file: Optional output file path (if None, returns bytes)
        
        Returns:
            PPTX file as bytes if output_file is None
        """
        # Create presentation
        prs = Presentation()
        
        # Set slide dimensions to 16:9 (width 10 inches, height 5.625 inches)
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(5.625)
        
        # Add each image as a slide
        for image_path in image_paths:
            if not os.path.exists(image_path):
                logger.warning(f"Image not found: {image_path}")
                continue
            
            # Add blank slide layout (layout 6 is typically blank)
            blank_slide_layout = prs.slide_layouts[6]
            slide = prs.slides.add_slide(blank_slide_layout)
            
            # Add image to fill entire slide
            slide.shapes.add_picture(
                image_path,
                left=0,
                top=0,
                width=prs.slide_width,
                height=prs.slide_height
            )
        
        # Save or return bytes
        if output_file:
            prs.save(output_file)
            return None
        else:
            # Save to bytes
            pptx_bytes = io.BytesIO()
            prs.save(pptx_bytes)
            pptx_bytes.seek(0)
            return pptx_bytes.getvalue()
    
    @staticmethod
    def create_pdf_from_images(image_paths: List[str], output_file: str = None) -> bytes:
        """
        Create PDF file from image paths
        
        Args:
            image_paths: List of absolute paths to images
            output_file: Optional output file path (if None, returns bytes)
        
        Returns:
            PDF file as bytes if output_file is None
        """
        images = []
        
        # Load all images
        for image_path in image_paths:
            if not os.path.exists(image_path):
                logger.warning(f"Image not found: {image_path}")
                continue
            
            img = Image.open(image_path)
            
            # Convert to RGB if necessary (PDF requires RGB)
            if img.mode != 'RGB':
                img = img.convert('RGB')
            
            images.append(img)
        
        if not images:
            raise ValueError("No valid images found for PDF export")
        
        # Save as PDF
        if output_file:
            images[0].save(
                output_file,
                save_all=True,
                append_images=images[1:],
                format='PDF'
            )
            return None
        else:
            # Save to bytes
            pdf_bytes = io.BytesIO()
            images[0].save(
                pdf_bytes,
                save_all=True,
                append_images=images[1:],
                format='PDF'
            )
            pdf_bytes.seek(0)
            return pdf_bytes.getvalue()

    @staticmethod
    def merge_pptx_files(pptx_paths: List[str], output_file: str = None) -> bytes:
        """
        Merge multiple PPTX files into one.
        Assumes all PPTX files are generated with compatible templates (default).
        """
        if not pptx_paths:
            raise ValueError("No PPTX files to merge")

        # Start with the first presentation as base
        # We assume the first one dictates the size/master, which works for our use case
        # where all are generated identically.
        base_prs = Presentation(pptx_paths[0])

        # Append subsequent presentations
        for path in pptx_paths[1:]:
            source_prs = Presentation(path)
            for source_slide in source_prs.slides:
                # Add a blank slide
                dest_slide = base_prs.slides.add_slide(base_prs.slide_layouts[6])

                # Copy shapes
                for shape in source_slide.shapes:
                    ExportService._copy_shape(shape, dest_slide)

        # Save
        if output_file:
            base_prs.save(output_file)
            return None
        else:
            out_bytes = io.BytesIO()
            base_prs.save(out_bytes)
            out_bytes.seek(0)
            return out_bytes.getvalue()

    @staticmethod
    def _copy_shape(shape, dest_slide):
        """Helper to copy a shape from one slide to another"""
        new_shape = None

        try:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                blob = shape.image.blob
                new_shape = dest_slide.shapes.add_picture(
                    io.BytesIO(blob), shape.left, shape.top, shape.width, shape.height
                )

            elif shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                new_shape = dest_slide.shapes.add_shape(
                    shape.auto_shape_type,
                    shape.left, shape.top, shape.width, shape.height
                )
                ExportService._copy_formatting(shape, new_shape)

            elif shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
                new_shape = dest_slide.shapes.add_textbox(
                    shape.left, shape.top, shape.width, shape.height
                )
                ExportService._copy_formatting(shape, new_shape)

            # Note: Other shape types like Groups, Tables, Charts are not currently supported
            # for the PPT Agent use case which primarily uses basic shapes.

        except Exception as e:
            logger.warning(f"Failed to copy shape {shape.shape_type}: {e}")

    @staticmethod
    def _copy_formatting(src_shape, new_shape):
        """Helper to copy formatting from source shape to new shape"""
        # Copy Fill
        try:
            if src_shape.fill.type:
                # Try to copy solid fill
                try:
                    # Accessing fore_color might raise TypeError if not solid/pattern
                    rgb = src_shape.fill.fore_color.rgb
                    new_shape.fill.solid()
                    new_shape.fill.fore_color.rgb = rgb
                except (TypeError, AttributeError):
                    pass
        except AttributeError:
            pass

        # Copy Line
        try:
            if src_shape.line.fill.type:
                try:
                    rgb = src_shape.line.color.rgb
                    new_shape.line.color.rgb = rgb
                    new_shape.line.width = src_shape.line.width
                except (TypeError, AttributeError):
                    pass
        except AttributeError:
            pass

        # Copy Text
        if src_shape.has_text_frame:
            new_shape.text_frame.clear()
            for paragraph in src_shape.text_frame.paragraphs:
                new_p = new_shape.text_frame.add_paragraph()
                new_p.alignment = paragraph.alignment

                for run in paragraph.runs:
                    new_r = new_p.add_run()
                    new_r.text = run.text
                    new_r.font.bold = run.font.bold
                    new_r.font.italic = run.font.italic
                    new_r.font.size = run.font.size
                    try:
                        new_r.font.color.rgb = run.font.color.rgb
                    except:
                        pass
                    new_r.font.name = run.font.name
