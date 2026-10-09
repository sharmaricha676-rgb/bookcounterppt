"""OOXML that python-pptx has no API for: slide transitions (fade / Morph / Page Curl) and an embedded 3D model.

Sources (see the research notes in the conversation): the Open XML SDK model3d schema and its
AnimatedModel3DExample sample for am3d; PowerPoint-saved captures for the p14/p15/p159 transition wrappers.
"""
import io
import random
import uuid

from lxml import etree
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
P15 = 'http://schemas.microsoft.com/office/powerpoint/2012/main'
P159 = 'http://schemas.microsoft.com/office/powerpoint/2015/09/main'
RT_MODEL3D = 'http://schemas.microsoft.com/office/2017/06/relationships/model3d'
CT_GLB = 'model/gltf.binary'  # what PowerPoint itself writes on save
EMU_IN = 914400


def _frag(xml):
    """parse a fragment that uses the p/a/r prefixes, return its single element"""
    decl = ' '.join(f'xmlns:{k}="{v}"' for k, v in NS.items())
    root = etree.fromstring(f'<wrap {decl}>{xml}</wrap>', etree.XMLParser(remove_blank_text=True))
    return root[0]


# ---------------------------------------------------------------- transitions
def _transition_xml(kind, dur):
    if kind == 'fade':
        return (f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice xmlns:p14="{P14}" Requires="p14">'
                f'<p:transition spd="med" p14:dur="{dur}"><p:fade/></p:transition></mc:Choice>'
                f'<mc:Fallback><p:transition spd="med"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>')
    if kind == 'morph':
        return (f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice xmlns:p159="{P159}" Requires="p159">'
                f'<p:transition spd="slow" xmlns:p14="{P14}" p14:dur="{dur}"><p159:morph option="byObject"/></p:transition>'
                f'</mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>')
    if kind == 'curl':
        return (f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice xmlns:p15="{P15}" Requires="p15">'
                f'<p:transition spd="slow" xmlns:p14="{P14}" p14:dur="{dur}"><p15:prstTrans prst="pageCurlDouble"/></p:transition>'
                f'</mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>')
    raise ValueError(kind)


def set_transition(slide, kind, dur_ms):
    """Put the transition in the p:transition slot: after p:clrMapOvr, before p:timing / p:extLst."""
    sld = slide._element
    for el in list(sld):
        if el.tag in (qn('p:transition'), f'{{{MC}}}AlternateContent'):
            sld.remove(el)
    el = _frag(_transition_xml(kind, dur_ms))
    anchor = sld.find(qn('p:clrMapOvr'))
    if anchor is None:
        anchor = sld.find(qn('p:cSld'))
    anchor.addnext(el)


# ---------------------------------------------------------------- 3D model
# rotation per view, in 60000ths of a degree (DrawingML angle units)
VIEWS = {
    'front': dict(ax=1080000, ay=1800000, az=0, fallback=0),       # 3/4 from the front right, a little from above
    'side': dict(ax=720000, ay=-3900000, az=0, fallback=29),       # from the left, showing the sensor arm
    'top': dict(ax=3900000, ay=600000, az=0, fallback=0),          # looking down at the screen
    'back': dict(ax=900000, ay=10800000 - 2100000, az=0, fallback=16),  # from behind
}
_model_part = {}

TEMPLATE = '''<mc:AlternateContent xmlns:mc="{MC}">
  <mc:Choice xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" Requires="am3d">
    <p:graphicFrame>
      <p:nvGraphicFramePr>
        <p:cNvPr id="{ID}" name="{NAME}" descr="{DESCR}">
          <a:extLst><a:ext uri="{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{GUID}"/></a:ext></a:extLst>
        </p:cNvPr>
        <p:cNvGraphicFramePr/>
        <p:nvPr><p:extLst><p:ext uri="{{D42A27DB-BD31-4B8C-83A1-F6EECF244321}}"><p14:modId xmlns:p14="{P14}" val="{MODID}"/></p:ext></p:extLst></p:nvPr>
      </p:nvGraphicFramePr>
      <p:xfrm><a:off x="{X}" y="{Y}"/><a:ext cx="{CX}" cy="{CY}"/></p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.microsoft.com/office/drawing/2017/model3d">
          <am3d:model3d r:embed="{MODEL_RID}">
            <am3d:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{CX}" cy="{CY}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></am3d:spPr>
            <am3d:camera>
              <am3d:pos x="0" y="0" z="67740115"/>
              <am3d:up dx="0" dy="36000000" dz="0"/>
              <am3d:lookAt x="0" y="0" z="0"/>
              <am3d:perspective fov="2700000"/>
            </am3d:camera>
            <am3d:trans>
              <am3d:meterPerModelUnit n="{MPU}" d="1000000"/>
              <am3d:preTrans dx="0" dy="0" dz="0"/>
              <am3d:scale><am3d:sx n="1000000" d="1000000"/><am3d:sy n="1000000" d="1000000"/><am3d:sz n="1000000" d="1000000"/></am3d:scale>
              <am3d:rot ax="{AX}" ay="{AY}" az="{AZ}"/>
              <am3d:postTrans dx="0" dy="0" dz="0"/>
            </am3d:trans>
            <am3d:raster rName="Office3DRenderer" rVer="16.0.8326"><am3d:blip r:embed="{IMG_RID}"/></am3d:raster>
            <am3d:objViewport viewportSz="{VP}"/>
            <am3d:ambientLight><am3d:clr><a:scrgbClr r="50000" g="50000" b="50000"/></am3d:clr><am3d:illuminance n="500000" d="1000000"/></am3d:ambientLight>
            <am3d:ptLight rad="0"><am3d:clr><a:scrgbClr r="100000" g="75000" b="50000"/></am3d:clr><am3d:intensity n="9765625" d="1000000"/><am3d:pos x="21959998" y="70920001" z="16344003"/></am3d:ptLight>
            <am3d:ptLight rad="0"><am3d:clr><a:scrgbClr r="40000" g="60000" b="95000"/></am3d:clr><am3d:intensity n="12250000" d="1000000"/><am3d:pos x="-37964106" y="51130435" z="57631972"/></am3d:ptLight>
            <am3d:ptLight rad="0"><am3d:clr><a:scrgbClr r="86837" g="72700" b="100000"/></am3d:clr><am3d:intensity n="3125000" d="1000000"/><am3d:pos x="-37739122" y="58056624" z="-34769649"/></am3d:ptLight>
          </am3d:model3d>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>
  </mc:Choice>
  <mc:Fallback>
    <p:pic>
      <p:nvPicPr>
        <p:cNvPr id="{ID}" name="{NAME}" descr="{DESCR}">
          <a:extLst><a:ext uri="{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{GUID}"/></a:ext></a:extLst>
        </p:cNvPr>
        <p:cNvPicPr><a:picLocks noGrp="1" noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1" noCrop="1"/></p:cNvPicPr>
        <p:nvPr/>
      </p:nvPicPr>
      <p:blipFill><a:blip r:embed="{IMG_RID}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="{X}" y="{Y}"/><a:ext cx="{CX}" cy="{CY}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
    </p:pic>
  </mc:Fallback>
</mc:AlternateContent>'''


def _next_id(slide):
    ids = [int(v) for v in slide.shapes._spTree.xpath('//@id') if str(v).isdigit()]
    return max(ids + [1]) + 1


def add_model3d(prs, slide, glb_path, view, work, assets, x, y, w, h, name='!!model',
                descr='3D model of the Book Counter bookmark'):
    """Embed the .glb as a real PowerPoint 3D model; apps without 3D support show the picture fallback.
    The .glb must be centred on the origin and in metres (tools/export-glb.cjs does both)."""
    pkg = prs.part.package
    if 'part' not in _model_part:
        with open(glb_path, 'rb') as f:
            _model_part['part'] = Part(PackURI('/ppt/media/model3d1.glb'), CT_GLB, pkg, f.read())
    model_rid = slide.part.relate_to(_model_part['part'], RT_MODEL3D)
    v = VIEWS[view]
    with open(f'{assets}/spin/{v["fallback"]:03d}.png', 'rb') as f:
        _, img_rid = slide.part.get_or_add_image_part(io.BytesIO(f.read()))

    X, Y, CX, CY = (int(round(t * EMU_IN)) for t in (x, y, w, h))
    xml = TEMPLATE.format(
        MC=MC, P14=P14, ID=_next_id(slide), NAME=name, DESCR=descr,
        GUID='{' + str(uuid.uuid4()).upper() + '}', MODID=random.randint(1, 4294967295),
        X=X, Y=Y, CX=CX, CY=CY, MODEL_RID=model_rid, IMG_RID=img_rid,
        MPU=4600000, AX=v['ax'], AY=v['ay'], AZ=v['az'], VP=int(round((CX ** 2 + CY ** 2) ** 0.5)))
    el = _frag(xml)
    tree = slide.shapes._spTree
    ext = tree.find(qn('p:extLst'))
    if ext is not None:
        ext.addprevious(el)
    else:
        tree.append(el)
    return el


def finalize(path):
    """Nothing to patch after save: python-pptx writes the .glb Override content type and every relationship."""
    return path
