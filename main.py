import glob
import os
import os.path
import shutil
import struct
import imghdr
from os.path import expanduser, join
import ctypes
import ctypes.wintypes

# All of the following class is made to enable coloring the forground console window in Windows.
# Based on http://stackoverflow.com/questions/17993814/why-the-irrelevant-code-made-a-difference
class ConsoleColor:
    if os.name is 'nt':

        # Windows constants.
        STD_OUTPUT_HANDLE = -11
        FOREGROUND_BLUE = 0x0001 # Text color contains blue.
        FOREGROUND_GREEN = 0x0002 # Text color contains green.
        FOREGROUND_RED = 0x0004 # Text color contains red.
        FOREGROUND_INTENSITY = 0x0008 # Text color is intensified.
        BACKGROUND_BLUE = 0x0010 # Background color contains blue.
        BACKGROUND_GREEN = 0x0020 # Background color contains green.
        BACKGROUND_RED = 0x0040 # Background color contains red.
        BACKGROUND_INTENSITY = 0x0080 # Background color is intensified.
        COMMON_LVB_LEADING_BYTE = 0x0100 # Leading byte.
        COMMON_LVB_TRAILING_BYTE = 0x0200 # Trailing byte.
        COMMON_LVB_GRID_HORIZONTAL = 0x0400 # Top horizontal
        COMMON_LVB_GRID_LVERTICAL = 0x0800 # Left vertical.
        COMMON_LVB_GRID_RVERTICAL = 0x1000 # Right vertical.
        COMMON_LVB_REVERSE_VIDEO = 0x4000 # Reverse foreground and background attribute.
        COMMON_LVB_UNDERSCORE = 0x8000 # Underscore.
        
        # Windows Structures
        class CONSOLE_SCREEN_BUFFER_INFO(ctypes.Structure):
            _fields_ = [
                ('dwSize', ctypes.wintypes._COORD),
                ('dwCursorPosition', ctypes.wintypes._COORD),
                ('wAttributes', ctypes.c_ushort),
                ('srWindow', ctypes.wintypes._SMALL_RECT),
                ('dwMaximumWindowSize', ctypes.wintypes._COORD)
            ]

        # Colors.
        ConsoleColorType = ctypes.c_ushort
        COLOR_MASK        = ConsoleColorType(0x000F)
        NEG_COLOR_MASK    = ConsoleColorType(0xFFF0)
        BLACK             = ConsoleColorType(0x0001)
        DARKBLUE          = ConsoleColorType(FOREGROUND_BLUE)
        DARKGREEN         = ConsoleColorType(FOREGROUND_GREEN)
        DARKCYAN          = ConsoleColorType(FOREGROUND_GREEN | FOREGROUND_BLUE)
        DARKRED           = ConsoleColorType(FOREGROUND_RED)
        DARKMAGENTA       = ConsoleColorType(FOREGROUND_RED | FOREGROUND_BLUE)
        DARKYELLOW        = ConsoleColorType(FOREGROUND_RED | FOREGROUND_GREEN)
        DARKGRAY          = ConsoleColorType(FOREGROUND_INTENSITY)
        GRAY              = ConsoleColorType(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE)
        BLUE              = ConsoleColorType(FOREGROUND_INTENSITY | FOREGROUND_BLUE)
        GREEN             = ConsoleColorType(FOREGROUND_INTENSITY | FOREGROUND_GREEN)
        CYAN              = ConsoleColorType(FOREGROUND_INTENSITY | FOREGROUND_GREEN | FOREGROUND_BLUE)
        RED               = ConsoleColorType(FOREGROUND_INTENSITY | FOREGROUND_RED)
        MAGENTA           = ConsoleColorType(FOREGROUND_INTENSITY | FOREGROUND_RED | FOREGROUND_BLUE)
        YELLOW            = ConsoleColorType(FOREGROUND_INTENSITY | FOREGROUND_RED | FOREGROUND_GREEN)
        WHITE             = ConsoleColorType(FOREGROUND_INTENSITY | FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE)

        # Declare the restype for GetStdHandle. This will mean that this code is ready to run under a 64 bit process.
        ctypes.windll.kernel32.GetStdHandle.restype = ctypes.wintypes.HANDLE

        # Static variables.
        hstd = ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)
    
        # A stack to hold the colors.
        textattribute_stack = []
        
        @staticmethod
        def set_foreground(color):
            # if type(color) is not ConsoleColorType:
            #   raise exeption('color should be of type ConsoleColorType failed.')
            
            csbi = ConsoleColor.CONSOLE_SCREEN_BUFFER_INFO()
            ret = ctypes.windll.kernel32.GetConsoleScreenBufferInfo(ConsoleColor.hstd, ctypes.byref(csbi))
            if not ret:
                raise exeption('GetConsoleScreenBufferInfo failed.')
            ConsoleColor.textattribute_stack.append(ConsoleColor.ConsoleColorType(csbi.wAttributes))
            
            textattribute = (ConsoleColor.textattribute_stack[-1].value & ConsoleColor.NEG_COLOR_MASK.value) | color.value
            
            ret = ctypes.windll.kernel32.SetConsoleTextAttribute(ConsoleColor.hstd, textattribute)
            if not ret:
                raise exeption('SetConsoleTextAttribute failed.')
                
        @staticmethod
        def pop_foreground():
            ret = ctypes.windll.kernel32.SetConsoleTextAttribute(ConsoleColor.hstd, ConsoleColor.textattribute_stack.pop())
            if not ret:
                raise exeption('SetConsoleTextAttribute failed.')
        

# http://stackoverflow.com/questions/8032642/how-to-obtain-image-size-using-standard-python-class-without-using-external-lib
def get_image_size(filename):
    """Determine the image type of file_handle and return its size. from draco

    :param filename: The image filename to get the dimensions of.
    :return: A tuple (width, height), or None if could not identify the image dimensions.
    """
    with open(filename, 'rb') as file_handle:
        head = file_handle.read(24)
        if len(head) != 24:
            return

        image_type = imghdr.what(filename)

        if image_type == 'png':
            check = struct.unpack('>i', head[4:8])[0]
            if check != 0x0d0a1a0a:
                return
            width, height = struct.unpack('>ii', head[16:24])
        elif image_type == 'gif':
            width, height = struct.unpack('<HH', head[6:10])
        elif image_type == 'jpeg':
            # noinspection PyBroadException
            try:
                file_handle.seek(0) # Read 0xff next
                size = 2
                file_type = 0
                while not 0xc0 <= file_type <= 0xcf:
                    file_handle.seek(size, 1)
                    byte = file_handle.read(1)
                    while ord(byte) == 0xff:
                        byte = file_handle.read(1)
                    file_type = ord(byte)
                    size = struct.unpack('>H', file_handle.read(2))[0] - 2
                # We are at a SOFn block
                file_handle.seek(1, 1)  # Skip `precision' byte.
                height, width = struct.unpack('>HH', file_handle.read(4))
            except Exception:
                return
        else:
            return
        return width, height


def pcolor(text, color=ConsoleColor.GREEN):
    """We use a little hack to allow this in midst of a print command - we return an empty string."""
    ConsoleColor.set_foreground(color)
    print text,
    ConsoleColor.pop_foreground()
    return ''


def merge(src_directory, src_filename_pattern, dst_directory, dst_filename_format, do_not_copy_filter=None):
    """
    :param src_directory: a source directory to copy the files from
    :param src_filename_pattern: a pattern that describes the filenames that should be copied from the directory
    :param dst_directory: a destination directory to copy the files to
    :param dst_filename_format:
    :param do_not_copy_filter:
    """
    
    for path_and_filename in glob.iglob(os.path.join(src_directory, src_filename_pattern)):
        title, ext = os.path.splitext(os.path.basename(path_and_filename))

        src_file = path_and_filename
        dst_file = os.path.join(dst_directory, dst_filename_format % title + ext)

        if do_not_copy_filter is not None and do_not_copy_filter(src_file):
            print pcolor('filtered file ', ConsoleColor.DARKYELLOW), src_file
            continue

        if os.path.isfile(dst_file):
            print pcolor('existing file ', ConsoleColor.DARKGRAY), dst_file
            continue

        print pcolor('copying file ', ConsoleColor.GREEN), src_file, pcolor(' to ', ConsoleColor.GREEN), dst_file
        shutil.copyfile(src_file, dst_file)


ASSETS_SRC_PATH_PREFIX = 'AppData\Local\Packages'
ASSETS_SRC_PATH_SUFIX  = 'LocalState\Assets'

LocalPackagesDirs = os.listdir(join(expanduser("~"), ASSETS_SRC_PATH_PREFIX))
ContentDeliveryManagerDirs = [dirname for dirname in LocalPackagesDirs if dirname.startswith('Microsoft.Windows.ContentDeliveryManager')]
if len(ContentDeliveryManagerDirs) is not 1:
    print 'I could not find the folder where Windows keeps its spotlight images. This calls for a programmer.'
    print 'Press any key to exit...', ; raw_input() ; exit()

ASSETS_SRC_PATH_RELATIVE = join(ASSETS_SRC_PATH_PREFIX, ContentDeliveryManagerDirs[0], ASSETS_SRC_PATH_SUFIX)
ASSETS_DST_PATH_RELATIVE = 'SpotlightAssets'

ASSETS_SRC_PATH = os.path.join(expanduser("~"), ASSETS_SRC_PATH_RELATIVE)
ASSETS_DST_PATH = os.path.join(expanduser("~"), ASSETS_DST_PATH_RELATIVE)


# Rename all the files in the temporary directory to the jpeg extension and move all assets to the destination
# directory. If the image file is too small, filter it.
merge(src_directory=ASSETS_SRC_PATH, src_filename_pattern=r'*',
      dst_directory=ASSETS_DST_PATH, dst_filename_format=r'%s.jpeg',
      do_not_copy_filter=lambda filename: get_image_size(filename) < (200, 200))
      
print 'Press any key to exit...', ; raw_input() ; exit()