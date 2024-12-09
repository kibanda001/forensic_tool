import argparse

from forensic_tool_meta import Forensic

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Forensic Tool")
    parser.add_argument("-pdf", "--pdf", dest="pdf", help="ADD PDF FILE PATH FOR ANALYSIS")
    parser.add_argument("-str", "--string", dest="string", help="ADD FILE THAT YOU CAN GET STRING")
    parser.add_argument("-docx", "--docx", dest="docx", help="ADD FILE WORD YOU CAN GET STRING")
    parser.add_argument("-xls", "--xls", dest="xls", help="ADD FILE Excel YOU CAN GET STRING")
    parser.add_argument("-img", "--image", dest="img", help="ADD IMAGE FILE")
    parser.add_argument("-gps", "--gps", dest="gps", help="GET IMAGE METADATA IF EXISTS")
    parser.add_argument("-cf", "--chrome", dest="cf", help="GET SITES VISITED BY USERS")
    parser.add_argument("-fc", "--fcookies", dest="fcookies", help="GET COOKIES OF SITES VISITED BY USERS")

    args = parser.parse_args()

    if args.pdf:
        Forensic.get_pdf_meta(args.pdf)
    if args.string:
        Forensic.get_pdf_meta(args.string)
    if args.docx:
        Forensic.get_docx_meta(args.docx)
        Forensic.get_docx_text(args.docx)
    if args.xls:
        Forensic.get_excel_meta(args.xls)
        Forensic.get_excel_text(args.xls)
    if args.img:
        Forensic.get_exif(args.img)
    if args.gps:
        Forensic.get_gps_from_exif(args.gps)
    if args.cf:
        Forensic.get_chrome_history(args.cf)

    if args.fcookies:
        Forensic.get_firefox_cookies(args.fcookies)
