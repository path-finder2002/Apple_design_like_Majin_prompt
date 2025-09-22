import json
from analyze_image import analyze
import sys

DPI = 144.0
PT_PER_PX = 72.0 / DPI  # convert pixel -> points


def px_to_pt(px):
    return px * PT_PER_PX


def summarize(path):
    data = analyze(path)
    width = data['width']
    height = data['height']
    segments = []
    for seg in data['segments']:
        # filter out obvious youtube overlays (thin band at top) and playback bar (bottom few px)
        if seg.height <= 6 and seg.width > width * 0.5:
            # playback bar
            continue
        if seg.start < 160:
            # youtube top chrome / captions
            continue
        if seg.pixel_count < 400:
            continue
        if seg.height < 12:
            # tiny noise
            continue
        segments.append({
            'top_px': seg.start,
            'bottom_px': seg.end,
            'height_px': seg.height,
            'height_pt': round(px_to_pt(seg.height), 1),
            'width_px': seg.width,
            'width_pt': round(px_to_pt(seg.width), 1),
            'min_x_px': seg.min_x,
            'max_x_px': seg.max_x,
            'max_row_px': seg.max_row_count,
            'pixel_count': seg.pixel_count,
        })
    return {
        'image': path,
        'width_px': width,
        'height_px': height,
        'segments': segments,
    }


def main(paths):
    results = [summarize(p) for p in paths]
    json.dump(results, sys.stdout, indent=2)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        raise SystemExit('Usage: summarize_images.py <image> [image...]')
    main(sys.argv[1:])
