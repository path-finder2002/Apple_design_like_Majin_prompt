import math
from collections import namedtuple
from png_utils import read_png

Segment = namedtuple('Segment', 'start end height min_x max_x width max_row_count pixel_count')


def analyze(path, bright_threshold=160, row_pixel_threshold=40, allow_gap=0, min_height=3):
    width, height, channels, pixels = read_png(path)
    # detect playback bar cutoff
    bar_cutoff = height
    for y in range(height - 1, -1, -1):
        found_red = False
        row_offset = y * width * channels
        for x in range(width):
            idx = row_offset + x * channels
            r = pixels[idx]
            g = pixels[idx + 1] if channels > 1 else r
            b = pixels[idx + 2] if channels > 2 else r
            a = pixels[idx + 3] if channels > 3 else 255
            if a < 50:
                continue
            if r > 180 and g < 120 and b < 120:
                found_red = True
                break
        if found_red:
            bar_cutoff = max(0, y - 5)
            break
    if bar_cutoff == height:
        bar_cutoff = height - 120
        if bar_cutoff < 0:
            bar_cutoff = height
    # compute row counts and detect segments
    row_counts = [0] * bar_cutoff
    for y in range(bar_cutoff):
        row_offset = y * width * channels
        count = 0
        for x in range(width):
            idx = row_offset + x * channels
            r = pixels[idx]
            g = pixels[idx + 1] if channels > 1 else r
            b = pixels[idx + 2] if channels > 2 else r
            a = pixels[idx + 3] if channels > 3 else 255
            if a < 50:
                continue
            if max(r, g, b) >= bright_threshold:
                count += 1
        row_counts[y] = count
    segments = []
    start = None
    gap = 0
    for y, count in enumerate(row_counts):
        if count > row_pixel_threshold:
            if start is None:
                start = y
            gap = 0
        else:
            if start is not None:
                if gap < allow_gap:
                    gap += 1
                else:
                    end = y - gap - 1
                    height_seg = end - start + 1
                    if height_seg >= min_height:
                        segments.append((start, end))
                    start = None
                    gap = 0
    if start is not None:
        end = len(row_counts) - 1
        height_seg = end - start + 1
        if height_seg >= min_height:
            segments.append((start, end))
    refined_segments = []
    for start, end in segments:
        min_x = width
        max_x = 0
        max_row = 0
        pixel_count = 0
        for y in range(start, end + 1):
            row_offset = y * width * channels
            row_count = 0
            local_min_x = width
            local_max_x = -1
            for x in range(width):
                idx = row_offset + x * channels
                r = pixels[idx]
                g = pixels[idx + 1] if channels > 1 else r
                b = pixels[idx + 2] if channels > 2 else r
                a = pixels[idx + 3] if channels > 3 else 255
                if a < 50:
                    continue
                if max(r, g, b) >= bright_threshold:
                    row_count += 1
                    pixel_count += 1
                    if x < local_min_x:
                        local_min_x = x
                    if x > local_max_x:
                        local_max_x = x
            if row_count > max_row:
                max_row = row_count
            if local_max_x >= local_min_x:
                if local_min_x < min_x:
                    min_x = local_min_x
                if local_max_x > max_x:
                    max_x = local_max_x
        width_seg = max_x - min_x + 1 if max_x >= min_x else 0
        refined_segments.append(Segment(start, end, end - start + 1, min_x, max_x, width_seg, max_row, pixel_count))
    return {
        'width': width,
        'height': height,
        'cutoff': bar_cutoff,
        'segments': refined_segments,
        'row_counts': row_counts,
    }


def main(paths):
    for path in paths:
        result = analyze(path)
        print(path)
        print('size:', result['width'], 'x', result['height'], 'cutoff:', result['cutoff'])
        for seg in result['segments']:
            print('  seg', seg.start, seg.end, 'height', seg.height, 'width', seg.width, 'min_x', seg.min_x, 'max_x', seg.max_x, 'max_row', seg.max_row_count, 'pixels', seg.pixel_count)

if __name__ == '__main__':
    import sys
    main(sys.argv[1:])
