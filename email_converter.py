#! /usr/bin/env python

import re
import logging
import os
import sys
import io
import base64

from functools import reduce

import email.message
import email.parser
import email.policy
from email.utils import parsedate_to_datetime, formatdate, formataddr

import compoundfiles
from rtfparse.parser import Rtf_Parser
from rtfparse.renderers.html_decapsulator import HTML_Decapsulator
import html2text

logger = logging.getLogger(__name__)

FALLBACK_ENCODING = "cp949"  # Changed to cp949 for better Korean support

# MAIN FUNCTIONS


def load(filename_or_stream):
    with compoundfiles.CompoundFileReader(filename_or_stream) as doc:
        doc.rtf_attachments = 0
        return load_message_stream(doc.root, True, doc)


def load_message_stream(entry, is_top_level, doc):
    # Load stream data.
    props = parse_properties(entry["__properties_version1.0"], is_top_level, entry, doc)

    # Construct the MIME message....
    msg = email.message.EmailMessage()

    # Add the raw headers, if known.
    if "TRANSPORT_MESSAGE_HEADERS" in props:
        # Get the string holding all of the headers.
        headers = props["TRANSPORT_MESSAGE_HEADERS"]
        if isinstance(headers, bytes):
            headers = headers.decode("utf-8")

        # Remove content-type header because the body we can get this
        # way is just the plain-text portion of the email and whatever
        # Content-Type header was in the original is not valid for
        # reconstructing it this way.
        headers = re.sub(r"Content-Type: .*(\n\s.*)*\n", "", headers, flags=re.I)

        # Parse them.
        headers = email.parser.HeaderParser(policy=email.policy.default).parsestr(
            headers
        )

        # Copy them into the message object.
        for header, value in headers.items():
            msg[header] = value

    else:
        # Construct common headers from metadata.
        if "MESSAGE_DELIVERY_TIME" in props:
            msg["Date"] = formatdate(props["MESSAGE_DELIVERY_TIME"].timestamp())
            del props["MESSAGE_DELIVERY_TIME"]

        if "SENDER_NAME" in props:
            if "SENT_REPRESENTING_NAME" in props:
                if props["SENT_REPRESENTING_NAME"]:
                    if props["SENDER_NAME"] != props["SENT_REPRESENTING_NAME"]:
                        props["SENDER_NAME"] += (
                            " (" + props["SENT_REPRESENTING_NAME"] + ")"
                        )
                del props["SENT_REPRESENTING_NAME"]
            if props["SENDER_NAME"]:
                msg["From"] = formataddr((props["SENDER_NAME"], ""))
            del props["SENDER_NAME"]

        if "DISPLAY_TO" in props:
            if props["DISPLAY_TO"]:
                msg["To"] = props["DISPLAY_TO"]
            del props["DISPLAY_TO"]

        if "DISPLAY_CC" in props:
            if props["DISPLAY_CC"]:
                msg["CC"] = props["DISPLAY_CC"]
            del props["DISPLAY_CC"]

        if "DISPLAY_BCC" in props:
            if props["DISPLAY_BCC"]:
                msg["BCC"] = props["DISPLAY_BCC"]
            del props["DISPLAY_BCC"]

        if "SUBJECT" in props:
            if props["SUBJECT"]:
                msg["Subject"] = decode_string(
                    props["SUBJECT"], [props.get("PR_INTERNET_CPID", FALLBACK_ENCODING)]
                )
            del props["SUBJECT"]

    # Add a plain text body from the BODY field.
    has_body = False
    if "BODY" in props:
        body = decode_string(
            props["BODY"], [props.get("PR_INTERNET_CPID", FALLBACK_ENCODING)]
        )
        msg.set_content(body, cte="quoted-printable")
        has_body = True

    # Add a HTML body from the RTF_COMPRESSED field.
    if "RTF_COMPRESSED" in props:
        # Decompress the value to Rich Text Format.
        import compressed_rtf

        rtf = props["RTF_COMPRESSED"]
        rtf = compressed_rtf.decompress(rtf)

        # Try rtfparse to de-encapsulate HTML stored in a rich
        # text container.
        try:
            rtf_blob = io.BytesIO(rtf)
            parsed = Rtf_Parser(rtf_file=rtf_blob).parse_file()
            html_stream = io.StringIO()
            HTML_Decapsulator().render(parsed, html_stream)
            html_body = html_stream.getvalue()

            if not has_body:
                # Try to convert that to plain/text if possible.
                text_body = html2text.html2text(html_body)
                msg.set_content(text_body, subtype="text", cte="quoted-printable")
                has_body = True

            if not has_body:
                msg.set_content(html_body, subtype="html", cte="quoted-printable")
                has_body = True
            else:
                msg.add_alternative(html_body, subtype="html", cte="quoted-printable")

        # If that fails, just attach the RTF file to the message.
        except:
            doc.rtf_attachments += 1
            fn = "messagebody_{}.rtf".format(doc.rtf_attachments)

            if not has_body:
                msg.set_content(
                    "<no plain text message body --- see attachment {}>".format(fn),
                    cte="quoted-printable",
                )
                has_body = True

            # Add RTF file as an attachment.
            msg.add_attachment(rtf, maintype="text", subtype="rtf", filename=fn)

    if not has_body:
        msg.set_content("<no message body>", cte="quoted-printable")

    # Add attachments.
    for stream in entry:
        if stream.name.startswith("__attach_version1.0_#"):
            try:
                process_attachment(msg, stream, doc)
            except KeyError as e:
                logger.error("Error processing attachment {} not found".format(str(e)))
                continue

    return msg


def process_attachment(msg, entry, doc):
    # Load attachment stream.
    props = parse_properties(entry["__properties_version1.0"], False, entry, doc)

    # The attachment content...
    blob = props["ATTACH_DATA_BIN"]

    # Get the filename and MIME type of the attachment.
    filename = (
        props.get("ATTACH_LONG_FILENAME")
        or props.get("ATTACH_FILENAME")
        or props.get("DISPLAY_NAME")
    )
    filename = decode_string(
        filename, [props.get("PR_INTERNET_CPID", FALLBACK_ENCODING)]
    )

    mime_type = props.get("ATTACH_MIME_TAG", "application/octet-stream")
    if isinstance(mime_type, bytes):
        mime_type = mime_type.decode("utf-8")

    filename = os.path.basename(filename)

    if isinstance(blob, bytes):
        # Handle image attachments
        if mime_type.startswith("image/"):
            img_data = base64.b64encode(blob).decode()
            msg.add_alternative(
                f'<img src="data:{mime_type};base64,{img_data}">', subtype="html"
            )

        msg.add_attachment(
            blob,
            maintype=mime_type.split("/", 1)[0],
            subtype=mime_type.split("/", 1)[-1],
            filename=filename,
        )
    elif isinstance(blob, str):
        msg.add_attachment(blob, filename=filename)
    else:  # a Message instance
        msg.add_attachment(blob, filename=filename)


def parse_properties(properties, is_top_level, container, doc):
    # Read a properties stream and return a Python dictionary
    # of the fields and values, using human-readable field names
    # in the mapping at the top of this module.

    # Load stream content.
    with doc.open(properties) as stream:
        stream = stream.read()

    # Skip header.
    i = 32 if is_top_level else 24

    # Read 16-byte entries.
    raw_properties = {}
    while i < len(stream):
        # Read the entry.
        property_type = stream[i + 0 : i + 2]
        property_tag = stream[i + 2 : i + 4]
        flags = stream[i + 4 : i + 8]
        value = stream[i + 8 : i + 16]
        i += 16

        # Turn the byte strings into numbers and look up the property type.
        property_type = property_type[0] + (property_type[1] << 8)
        property_tag = property_tag[0] + (property_tag[1] << 8)
        if property_tag not in property_tags:
            continue  # should not happen
        tag_name, _ = property_tags[property_tag]
        tag_type = property_types.get(property_type)

        # Fixed Length Properties.
        if isinstance(tag_type, FixedLengthValueLoader):
            # The value comes from the stream above.
            pass

        # Variable Length Properties.
        elif isinstance(tag_type, VariableLengthValueLoader):
            value_length = stream[i + 8 : i + 12]  # not used

            # Look up the stream in the document that holds the value.
            streamname = "__substg1.0_{0:0{1}X}{2:0{3}X}".format(
                property_tag, 4, property_type, 4
            )
            try:
                with doc.open(container[streamname]) as innerstream:
                    value = innerstream.read()
            except:
                # Stream isn't present!
                logger.error("stream missing {}".format(streamname))
                continue

        elif isinstance(tag_type, EMBEDDED_MESSAGE):
            # Look up the stream in the document that holds the attachment.
            streamname = "__substg1.0_{0:0{1}X}{2:0{3}X}".format(
                property_tag, 4, property_type, 4
            )
            try:
                value = container[streamname]
            except:
                # Stream isn't present!
                logger.error("stream missing {}".format(streamname))
                continue

        else:
            # unrecognized type
            logger.error("unhandled property type {}".format(hex(property_type)))
            continue

        raw_properties[tag_name] = (tag_type, value)

    # Decode all FixedLengthValueLoader properties so we have codepage
    # properties.
    properties = {}
    for tag_name, (tag_type, value) in raw_properties.items():
        if not isinstance(tag_type, FixedLengthValueLoader):
            continue
        try:
            properties[tag_name] = tag_type.load(value)
        except Exception as e:
            logger.error("Error while reading stream: {}".format(str(e)))

    # String8 strings use code page information stored in other
    # properties, which may not be present. Find the Python
    # encoding to use.

    # The encoding of the "BODY" (and HTML body) properties.
    body_encoding = None
    if (
        "PR_INTERNET_CPID" in properties
        and properties["PR_INTERNET_CPID"] in code_pages
    ):
        body_encoding = code_pages[properties["PR_INTERNET_CPID"]]

    # The encoding of "string properties of the message object".
    properties_encoding = None
    if (
        "PR_MESSAGE_CODEPAGE" in properties
        and properties["PR_MESSAGE_CODEPAGE"] in code_pages
    ):
        properties_encoding = code_pages[properties["PR_MESSAGE_CODEPAGE"]]

    # Decode all of the remaining properties.
    for tag_name, (tag_type, value) in raw_properties.items():
        if isinstance(tag_type, FixedLengthValueLoader):
            continue  # already done, above

        # The codepage properties may be wrong. Fall back to
        # the other property if present.
        encodings = (
            [body_encoding, properties_encoding]
            if tag_name == "BODY"
            else [properties_encoding, body_encoding]
        )

        try:
            properties[tag_name] = tag_type.load(value, encodings=encodings, doc=doc)
        except KeyError as e:
            logger.error("Error while reading stream: {} not found".format(str(e)))
        except Exception as e:
            logger.error("Error while reading stream: {}".format(str(e)))

    return properties


def decode_string(value, encodings):
    if isinstance(value, bytes):
        for encoding in encodings + [FALLBACK_ENCODING, "utf-8"]:
            try:
                return value.decode(encoding)
            except UnicodeDecodeError:
                continue
    return value


# PROPERTY VALUE LOADERS


class FixedLengthValueLoader(object):
    pass


class NULL(FixedLengthValueLoader):
    @staticmethod
    def load(value):
        # value is an eight-byte long bytestring with unused content.
        return None


class BOOLEAN(FixedLengthValueLoader):
    @staticmethod
    def load(value):
        # value is an eight-byte long bytestring holding a two-byte integer.
        return value[0] == 1


class INTEGER16(FixedLengthValueLoader):
    @staticmethod
    def load(value):
        # value is an eight-byte long bytestring holding a two-byte integer.
        return reduce(lambda a, b: (a << 8) + b, reversed(value[0:2]))


class INTEGER32(FixedLengthValueLoader):
    @staticmethod
    def load(value):
        # value is an eight-byte long bytestring holding a four-byte integer.
        return reduce(lambda a, b: (a << 8) + b, reversed(value[0:4]))


class INTEGER64(FixedLengthValueLoader):
    @staticmethod
    def load(value):
        # value is an eight-byte long bytestring holding an eight-byte integer.
        return reduce(lambda a, b: (a << 8) + b, reversed(value))


class INTTIME(FixedLengthValueLoader):
    @staticmethod
    def load(value):
        # value is an eight-byte long bytestring encoding the integer number of
        # 100-nanosecond intervals since January 1, 1601.
        from datetime import datetime, timedelta

        value = reduce(
            lambda a, b: (a << 8) + b, reversed(value)
        )  # bytestring to integer
        try:
            value = datetime(1601, 1, 1) + timedelta(seconds=value / 10000000)
        except OverflowError:
            value = None
        return value


class VariableLengthValueLoader(object):
    pass


class BINARY(VariableLengthValueLoader):
    @staticmethod
    def load(value, **kwargs):
        # value is a bytestring. Just return it.
        return value


class STRING8(VariableLengthValueLoader):
    @staticmethod
    def load(value, encodings, **kwargs):
        # Value is a "bytestring" and encodings is a list of Python
        # codecs to try. If all fail, try the fallback codec with
        # character replacement so that this never fails.
        return decode_
