import struct
import zlib

def read_png(path):
    with open(path, 'rb') as f:
        signature = f.read(8)
        if signature != b'\x89PNG\r\n\x1a\n':
            raise ValueError('Not a PNG file')
        width = height = None
        bit_depth = color_type = None
        data = bytearray()
        while True:
            chunk_len_bytes = f.read(4)
            if not chunk_len_bytes:
                break
            chunk_len = struct.unpack('>I', chunk_len_bytes)[0]
            chunk_type = f.read(4)
            chunk_data = f.read(chunk_len)
            f.read(4)  # CRC, ignore
            if chunk_type == b'IHDR':
                width, height, bit_depth, color_type, compression, filter_method, interlace = struct.unpack('>IIBBBBB', chunk_data)
                if bit_depth != 8:
                    raise NotImplementedError('Only 8-bit depth supported')
                if color_type not in (2, 6, 0, 3):
                    raise NotImplementedError(f'Unsupported color type {color_type}')
                if interlace != 0:
                    raise NotImplementedError('Interlaced PNG not supported')
            elif chunk_type == b'IDAT':
                data.extend(chunk_data)
            elif chunk_type == b'IEND':
                break
        if width is None or height is None:
            raise ValueError('Missing IHDR chunk')
        decompressed = zlib.decompress(data)
        if color_type == 6:
            channels = 4
        elif color_type == 2:
            channels = 3
        elif color_type == 0:
            channels = 1
        elif color_type == 3:
            raise NotImplementedError('Indexed color not supported')
        bytes_per_pixel = channels
        stride = width * channels
        expected_len = (stride + 1) * height
        if len(decompressed) != expected_len:
            raise ValueError('Unexpected decompressed data length')
        pixels = bytearray(height * stride)
        src_offset = 0
        dest_offset = 0
        prev_scanline = bytearray(stride)
        for _ in range(height):
            filter_type = decompressed[src_offset]
            src_offset += 1
            scanline = decompressed[src_offset:src_offset + stride]
            src_offset += stride
            recon = bytearray(stride)
            if filter_type == 0:  # None
                recon[:] = scanline
            elif filter_type == 1:  # Sub
                for i in range(stride):
                    left = recon[i - bytes_per_pixel] if i >= bytes_per_pixel else 0
                    recon[i] = (scanline[i] + left) & 0xFF
            elif filter_type == 2:  # Up
                for i in range(stride):
                    recon[i] = (scanline[i] + prev_scanline[i]) & 0xFF
            elif filter_type == 3:  # Average
                for i in range(stride):
                    left = recon[i - bytes_per_pixel] if i >= bytes_per_pixel else 0
                    up = prev_scanline[i]
                    recon[i] = (scanline[i] + ((left + up) >> 1)) & 0xFF
            elif filter_type == 4:  # Paeth
                def paeth_predictor(a, b, c):
                    p = a + b - c
                    pa = abs(p - a)
                    pb = abs(p - b)
                    pc = abs(p - c)
                    if pa <= pb and pa <= pc:
                        return a
                    elif pb <= pc:
                        return b
                    else:
                        return c
                for i in range(stride):
                    left = recon[i - bytes_per_pixel] if i >= bytes_per_pixel else 0
                    up = prev_scanline[i]
                    up_left = prev_scanline[i - bytes_per_pixel] if i >= bytes_per_pixel else 0
                    recon[i] = (scanline[i] + paeth_predictor(left, up, up_left)) & 0xFF
            else:
                raise NotImplementedError(f'Unsupported filter type {filter_type}')
            pixels[dest_offset:dest_offset + stride] = recon
            prev_scanline[:] = recon
            dest_offset += stride
        return width, height, channels, pixels

if __name__ == '__main__':
    import sys
    w, h, c, pixels = read_png(sys.argv[1])
    print('Loaded', sys.argv[1], '->', w, h, c)
