"""Namespace-preserving XML parse/serialize for OOXML parts.

Plain ElementTree rewrites `p:` / `a:` prefixes to `ns0:` / `ns1:` on output and drops
declarations it thinks are unused. PowerPoint may repair that; LibreOffice refuses to load it.
Every script that rewrites a part must go through parse()/serialize() here.
"""
import re
import xml.etree.ElementTree as ET

_DECL = re.compile(rb'xmlns(?::([A-Za-z_][\w.-]*))?="([^"]*)"')
_IGNORABLE = re.compile(rb'Ignorable="([^"]*)"')
_PREFIX_URI = {}   # prefix -> uri, everything ever seen
_DEFAULT_URIS = set()  # uris that appeared as a default (unprefixed) namespace

# Standard OOXML prefixes, registered up front so output is stable even for parts we never parsed.
_KNOWN = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main',
    'p15': 'http://schemas.microsoft.com/office/powerpoint/2012/main',
    'a14': 'http://schemas.microsoft.com/office/drawing/2010/main',
    'a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
    'v': 'urn:schemas-microsoft-com:vml',
    'o': 'urn:schemas-microsoft-com:office:office',
    'vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'dcterms': 'http://purl.org/dc/terms/',
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
}
for _p, _u in _KNOWN.items():
    ET.register_namespace(_p, _u); _PREFIX_URI[_p] = _u


def parse(data):
    """ET.fromstring, but remember every xmlns declaration so serialize() can reproduce it."""
    if isinstance(data, str):
        data = data.encode('utf-8')
    for prefix, uri in _DECL.findall(data):
        p, u = prefix.decode(), uri.decode()
        if not p:
            _DEFAULT_URIS.add(u)
        elif not re.fullmatch(r'ns\d+', p):
            if _PREFIX_URI.get(p) != u:
                ET.register_namespace(p, u); _PREFIX_URI[p] = u
    return ET.fromstring(data)


def serialize(elem):
    """ET.tostring with original prefixes, a real default namespace for rels/content-types,
    and every prefix named in mc:Ignorable declared on the root."""
    out = ET.tostring(elem, encoding='utf-8', xml_declaration=True)
    # Unregistered URIs come back as nsN. That only happens for default-namespace documents
    # (Relationships, Types, core properties) — turn them back into a default namespace.
    for m in re.finditer(rb'xmlns:(ns\d+)="([^"]*)"', out):
        pfx, uri = m.group(1), m.group(2)
        if uri.decode() in _DEFAULT_URIS:
            out = out.replace(b'xmlns:' + pfx + b'="', b'xmlns="', 1)
            out = out.replace(b'<' + pfx + b':', b'<').replace(b'</' + pfx + b':', b'</')
    # Re-declare prefixes listed in mc:Ignorable that ET dropped because no element used them.
    head_end = out.find(b'>', out.find(b'<', out.find(b'?>') + 2))
    head = out[:head_end]
    ign = _IGNORABLE.search(head)
    if ign:
        missing = b''
        for pfx in ign.group(1).split():
            p = pfx.decode()
            if b'xmlns:' + pfx + b'="' not in head and p in _PREFIX_URI:
                missing += b' xmlns:' + pfx + b'="' + _PREFIX_URI[p].encode() + b'"'
        if missing:
            out = head + missing + out[head_end:]
    return out
