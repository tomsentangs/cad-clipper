"""
CAD裁剪服务 - Flask HTTP接口
提供CAD文件裁剪、进度查询、文件下载功能
支持DXF和DWG格式
"""

import os
import sys
import uuid
import threading
from datetime import datetime
# 1. 检查是否是打包后的exe
if getattr(sys, 'frozen', False):
    # 运行打包后的exe
    application_path = sys._MEIPASS
    
    # 设置GDAL环境变量,使用内置的GDAL数据
    os.environ['GDAL_DATA'] = os.path.join(application_path, 'gdal_data')
    os.environ['PROJ_LIB'] = os.path.join(application_path, 'proj_data')
    os.environ['GDAL_DRIVER_PATH'] = 'disabled'  # 禁用插件路径
    
    print(f"[打包模式] 使用内置GDAL环境")
    print(f"[GDAL_DATA] {os.environ['GDAL_DATA']}")
    print(f"[PROJ_LIB] {os.environ['PROJ_LIB']}")
else:
    # 开发模式
    print(f"[开发模式] 使用系统GDAL环境")

# 2. 禁用GDAL的调试输出 (这是关键!)
os.environ['CPL_DEBUG'] = 'OFF'
os.environ['CPL_LOG'] = 'NUL'  # Windows下重定向到NUL
os.environ['CPL_CURL_VERBOSE'] = 'NO'

# 3. 现在导入GDAL相关库
try:
    from osgeo import gdal
    # 进一步禁用GDAL的日志输出
    gdal.SetConfigOption('CPL_DEBUG', 'OFF')
    gdal.SetConfigOption('CPL_LOG', 'NUL')
    gdal.UseExceptions()  # 使用异常而不是错误代码
    print(f"[✓] GDAL {gdal.__version__} 加载成功")
except Exception as e:
    print(f"[✗] GDAL加载失败: {e}")
    sys.exit(1)
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import ezdxf
from ezdxf.addons import odafc
import geopandas as gpd
from shapely.geometry import Polygon, Point, LineString, MultiPolygon
from shapely.ops import unary_union
from shapely.geometry import MultiLineString, GeometryCollection
import logging
import argparse


# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# 创建Flask应用
app = Flask(__name__)
CORS(app)  # 允许跨域请求

# 任务状态存储
tasks = {}
task_lock = threading.Lock()


def is_dwg_file(file_path):
    """判断是否为DWG文件"""
    return file_path.lower().endswith('.dwg')


def is_dxf_file(file_path):
    """判断是否为DXF文件"""
    return file_path.lower().endswith('.dxf')


def load_cad_file(file_path):
    """
    加载CAD文件,支持DXF和DWG格式
    
    Args:
        file_path: CAD文件路径
        
    Returns:
        doc: ezdxf文档对象
    """
    try:
        if is_dwg_file(file_path):
            logger.info(f"检测到DWG文件,正在转换...")
            # DWG需要先转换为DXF
            # 方法1: 使用ODA File Converter (需要安装)
            try:
                # 尝试使用ezdxf的ODA转换器
                temp_dxf = file_path + '.temp.dxf'
                odafc.export_dwg(file_path, temp_dxf)
                doc = ezdxf.readfile(temp_dxf)
                # 删除临时文件
                if os.path.exists(temp_dxf):
                    os.remove(temp_dxf)
                logger.info("DWG文件转换成功")
                return doc
            except Exception as e:
                logger.warning(f"ODA转换失败,尝试直接读取: {e}")
                # 方法2: 直接读取DWG (ezdxf 0.17+支持)
                doc = ezdxf.readfile(file_path)
                logger.info("直接读取DWG文件成功")
                return doc
        else:
            # DXF文件直接读取
            doc = ezdxf.readfile(file_path)
            logger.info("读取DXF文件成功")
            return doc
            
    except Exception as e:
        logger.error(f"加载CAD文件失败: {e}")
        raise Exception(f"无法读取CAD文件: {str(e)}")


def save_cad_file(doc, output_path):
    """
    保存CAD文件,根据扩展名自动选择格式
    
    Args:
        doc: ezdxf文档对象
        output_path: 输出文件路径
    """
    try:
        if is_dwg_file(output_path):
            logger.info("保存为DWG格式...")
            # 先保存为DXF
            temp_dxf = output_path + '.temp.dxf'
            doc.saveas(temp_dxf)
            
            # 转换为DWG
            try:
                odafc.export_dwg(temp_dxf, output_path)
                # 删除临时文件
                if os.path.exists(temp_dxf):
                    os.remove(temp_dxf)
                logger.info("DWG文件保存成功")
            except Exception as e:
                logger.warning(f"无法保存为DWG格式: {e}")
                logger.info("将保存为DXF格式代替")
                # 如果转换失败,保存为DXF
                output_path = output_path.rsplit('.', 1)[0] + '.dxf'
                doc.saveas(output_path)
        else:
            # 保存为DXF
            doc.saveas(output_path)
            logger.info("DXF文件保存成功")
            
    except Exception as e:
        logger.error(f"保存CAD文件失败: {e}")
        raise Exception(f"无法保存CAD文件: {str(e)}")


class ClipTask:
    """裁剪任务类"""
    def __init__(self, task_id, input_dxf, shp_path, output_path):
        self.task_id = task_id
        self.input_dxf = input_dxf
        self.shp_path = shp_path
        self.output_path = output_path
        self.status = 'pending'  # pending, processing, completed, failed
        self.progress = 0
        self.total = 0
        self.message = '等待开始'
        self.error = None
        self.start_time = datetime.now()
        self.end_time = None


def build_default_output_path(input_path):
    base, ext = os.path.splitext(input_path)
    if ext.lower() not in ('.dxf', '.dwg'):
        ext = '.dxf'
    return f"{base}_clip{ext}"


def load_shp(shp_path):
    """加载SHP文件"""
    try:
        gdf = gpd.read_file(shp_path)
        logger.info(f"读取SHP文件: {shp_path}, {len(gdf)}个要素")
        
        polygons = []
        for geom in gdf.geometry:
            if geom is None or geom.is_empty:
                continue
            if isinstance(geom, Polygon):
                polygons.append(geom)
            elif isinstance(geom, MultiPolygon):
                polygons.extend(list(geom.geoms))
        
        if not polygons:
            return None
            
        merged_geom = unary_union(polygons)
        
        if isinstance(merged_geom, (Polygon, MultiPolygon)):
            merged_polygon = merged_geom
        elif isinstance(merged_geom, GeometryCollection):
            polys = [g for g in merged_geom.geoms if isinstance(g, (Polygon, MultiPolygon))]
            if polys:
                merged_polygon = unary_union(polys)
            else:
                return None
        else:
            return None
        
        if not merged_polygon.is_valid:
            merged_polygon = merged_polygon.buffer(0)
        
        return merged_polygon
        
    except Exception as e:
        logger.error(f"读取SHP文件失败: {e}")
        return None


def clip_line_entity(entity, shp_polygon, msp):
    """裁剪LINE实体"""
    to_delete = []
    to_add = []
    
    try:
        if not isinstance(shp_polygon, (Polygon, MultiPolygon)):
            to_delete.append(entity)
            return to_delete, to_add
            
        start_point = Point(entity.dxf.start.x, entity.dxf.start.y)
        end_point = Point(entity.dxf.end.x, entity.dxf.end.y)
        line = LineString([start_point, end_point])

        if shp_polygon.contains(line):
            return to_delete, to_add
            
        clipped = line.intersection(shp_polygon)
        
        if clipped.is_empty or isinstance(clipped, Point):
            to_delete.append(entity)
        elif isinstance(clipped, LineString):
            coords = list(clipped.coords)
            if len(coords) >= 2:
                to_delete.append(entity)
                new_line = msp.add_line(
                    start=(coords[0][0], coords[0][1]),
                    end=(coords[-1][0], coords[-1][1]),
                    dxfattribs={
                        'layer': entity.dxf.layer,
                        'color': entity.dxf.color if hasattr(entity.dxf, 'color') else None,
                        'linetype': entity.dxf.linetype if hasattr(entity.dxf, 'linetype') else None,
                    }
                )
                to_add.append(new_line)
            else:
                to_delete.append(entity)
        elif isinstance(clipped, MultiLineString):
            to_delete.append(entity)
            for subline in clipped.geoms:
                coords = list(subline.coords)
                if len(coords) >= 2:
                    new_line = msp.add_line(
                        start=(coords[0][0], coords[0][1]),
                        end=(coords[-1][0], coords[-1][1]),
                        dxfattribs={
                            'layer': entity.dxf.layer,
                            'color': entity.dxf.color if hasattr(entity.dxf, 'color') else None,
                            'linetype': entity.dxf.linetype if hasattr(entity.dxf, 'linetype') else None,
                        }
                    )
                    to_add.append(new_line)
        elif isinstance(clipped, GeometryCollection):
            to_delete.append(entity)
            for geom in clipped.geoms:
                if isinstance(geom, LineString):
                    coords = list(geom.coords)
                    if len(coords) >= 2:
                        new_line = msp.add_line(
                            start=(coords[0][0], coords[0][1]),
                            end=(coords[-1][0], coords[-1][1]),
                            dxfattribs={
                                'layer': entity.dxf.layer,
                                'color': entity.dxf.color if hasattr(entity.dxf, 'color') else None,
                                'linetype': entity.dxf.linetype if hasattr(entity.dxf, 'linetype') else None,
                            }
                        )
                        to_add.append(new_line)
        else:
            to_delete.append(entity)
            
    except Exception as e:
        to_delete.append(entity)
    
    return to_delete, to_add


def clip_lwpolyline_entity(entity, shp_polygon, msp):
    """裁剪LWPOLYLINE实体"""
    to_delete = []
    to_add = []
    
    try:
        if not isinstance(shp_polygon, (Polygon, MultiPolygon)):
            to_delete.append(entity)
            return to_delete, to_add
            
        points = [(p[0], p[1]) for p in entity.get_points('xy')]
        
        if len(points) < 2:
            to_delete.append(entity)
            return to_delete, to_add
        
        if entity.closed and points[0] != points[-1]:
            points.append(points[0])
        
        line = LineString(points)
        
        if shp_polygon.contains(line):
            return to_delete, to_add
        
        clipped = line.intersection(shp_polygon)
        
        to_delete.append(entity)
        
        if clipped.is_empty:
            pass
        elif isinstance(clipped, LineString):
            coords = list(clipped.coords)
            new_poly = msp.add_lwpolyline(
                coords,
                dxfattribs={
                    'layer': entity.dxf.layer,
                    'color': entity.dxf.color if hasattr(entity.dxf, 'color') else None,
                    'linetype': entity.dxf.linetype if hasattr(entity.dxf, 'linetype') else None,
                }
            )
            to_add.append(new_poly)
        elif isinstance(clipped, MultiLineString):
            for subline in clipped.geoms:
                coords = list(subline.coords)
                if len(coords) >= 2:
                    if len(coords) == 2:
                        new_line = msp.add_line(
                            start=coords[0],
                            end=coords[1],
                            dxfattribs={'layer': entity.dxf.layer}
                        )
                        to_add.append(new_line)
                    else:
                        new_poly = msp.add_lwpolyline(
                            coords,
                            dxfattribs={'layer': entity.dxf.layer}
                        )
                        to_add.append(new_poly)
                        
    except Exception as e:
        to_delete.append(entity)
    
    return to_delete, to_add


def clip_circle_entity(entity, shp_polygon, msp):
    """裁剪CIRCLE实体"""
    to_delete = []
    to_add = []
    
    try:
        if not isinstance(shp_polygon, (Polygon, MultiPolygon)):
            to_delete.append(entity)
            return to_delete, to_add
            
        center = Point(entity.dxf.center.x, entity.dxf.center.y)
        radius = entity.dxf.radius
        circle_geom = center.buffer(radius)
        
        if not circle_geom.intersects(shp_polygon):
            to_delete.append(entity)
            
    except Exception as e:
        to_delete.append(entity)
    
    return to_delete, to_add


def clip_arc_entity(entity, shp_polygon, msp):
    """裁剪ARC实体"""
    to_delete = []
    to_add = []
    
    try:
        if not isinstance(shp_polygon, (Polygon, MultiPolygon)):
            to_delete.append(entity)
            return to_delete, to_add
            
        center = Point(entity.dxf.center.x, entity.dxf.center.y)
        radius = entity.dxf.radius
        arc_geom = center.buffer(radius)
        
        if not arc_geom.intersects(shp_polygon):
            to_delete.append(entity)
            
    except Exception as e:
        to_delete.append(entity)
    
    return to_delete, to_add


def clip_text_entity(entity, shp_polygon, msp):
    """裁剪TEXT/MTEXT实体"""
    to_delete = []
    to_add = []
    
    try:
        if not isinstance(shp_polygon, (Polygon, MultiPolygon)):
            to_delete.append(entity)
            return to_delete, to_add
        
        insert_point = Point(entity.dxf.insert.x, entity.dxf.insert.y)
        
        if not shp_polygon.contains(insert_point):
            to_delete.append(entity)
        #else:
        #    to_add.append(entity)        
    except Exception as e:
        to_delete.append(entity)
    
    return to_delete, to_add


from shapely.geometry import Polygon, MultiPolygon, LineString, Point

def clip_hatch_entity(entity, shp_polygon, msp):
    """裁剪HATCH实体"""
    to_delete = []
    to_add = []
    
    try:
        # 如果shp_polygon不是Polygon或MultiPolygon类型，则跳过
        if not isinstance(shp_polygon, (Polygon, MultiPolygon)):
            to_delete.append(entity)
            return to_delete, to_add
        
        has_intersection = False
        
        for path in entity.paths:
            if hasattr(path, 'vertices'):
                vertices = [(v[0], v[1]) for v in path.vertices]
                
                # 如果路径包含多个顶点
                if len(vertices) >= 2:
                    if len(vertices) > 2 and path.path_type_flags & 2:
                        if vertices[0] != vertices[-1]:
                            vertices.append(vertices[0])
                        hatch_geom = Polygon(vertices)  # 使用Polygon表示HATCH
                    else:
                        hatch_geom = LineString(vertices)  # 使用LineString表示HATCH
                    
                    # 检查是否与shp_polygon相交
                    if hatch_geom.intersects(shp_polygon):
                        has_intersection = True
                        #intersection_geom = hatch_geom.intersection(shp_polygon)                     
                        # 如果相交，保留相交部分
                        #if intersection_geom.is_valid and not intersection_geom.is_empty:
                            # 将相交部分添加到to_add中
                            #new_vertices = list(intersection_geom.exterior.coords) if isinstance(intersection_geom, Polygon) else list(intersection_geom.coords)
                            #path.vertices = new_vertices  # 更新路径的顶点                         
                        #else:
                        #    to_delete.append(entity)  # 如果无有效交集，则删除
                        break

            elif hasattr(path, 'edges'):
                # 处理边缘类型
                for edge in path.edges:
                    if edge.EDGE_TYPE == "LineEdge":
                        start = Point(edge.start[0], edge.start[1])
                        if shp_polygon.contains(start):
                            has_intersection = True
                            break
                    elif edge.EDGE_TYPE in ["ArcEdge", "EllipseEdge"]:
                        center = Point(edge.center[0], edge.center[1])
                        if shp_polygon.contains(center):
                            has_intersection = True
                            break
            
            if has_intersection:
                break
        
        if not has_intersection:
            to_delete.append(entity)
    
    except Exception as e:
        pass
    
    return to_delete, to_add





def clip_point_entity(entity, shp_polygon, msp):
    """裁剪POINT实体"""
    to_delete = []
    to_add = []
    
    try:
        if not isinstance(shp_polygon, (Polygon, MultiPolygon)):
            to_delete.append(entity)
            return to_delete, to_add
        
        point = Point(entity.dxf.location.x, entity.dxf.location.y)
        
        if not shp_polygon.contains(point):
            to_delete.append(entity)
            
    except Exception as e:
        to_delete.append(entity)
    
    return to_delete, to_add

def clip_block_entity(block_entity, shp_polygon, msp):
    to_delete = []
    to_add = []
    
    # 遍历块中的所有子实体
    for sub_entity in block_entity.get_sub_entities():
        if sub_entity.EntityType == 'Hatch':
            sub_to_delete, sub_to_add = clip_hatch_entity(sub_entity, shp_polygon, msp)
            to_delete.extend(sub_to_delete)
            to_add.extend(sub_to_add)
        # 可以添加更多的类型判断，例如 Line, Circle 等
    return to_delete, to_add


def process_clip_task(task_id):
    """执行裁剪任务"""
    with task_lock:
        task = tasks.get(task_id)
        if not task:
            return
        task.status = 'processing'
        task.message = '正在处理...'
    
    try:
        # 1. 加载SHP文件
        with task_lock:
            task.message = '正在读取SHP文件...'
            task.progress = 5
        
        shp_polygon = load_shp(task.shp_path)
        if shp_polygon is None:
            raise Exception("无法读取SHP文件或提取有效多边形")
        
        # 2. 读取DXF文件
        with task_lock:
            task.message = '正在读取DXF文件...'
            task.progress = 10
        
        doc = load_cad_file(task.input_dxf)
        print("开始加载cad")
        #doc = load_cad_file(task.input_dxf)
        msp = doc.modelspace()
        entity_list = list(msp)
        entity_count = len(entity_list)
        
        with task_lock:
            task.total = entity_count
            task.message = f'开始裁剪 {entity_count} 个实体...'
            task.progress = 15
        
        # 3. 裁剪实体
        all_to_delete = []
        all_to_add = []
        
        if entity_count == 0:
            with task_lock:
                task.message = 'DXF文件中没有可处理实体...'
                task.progress = 85
        else:
            for idx, entity in enumerate(entity_list):
                entity_type = entity.dxftype()
                to_delete, to_add = [], []
                
                if entity_type == 'LINE':
                    to_delete, to_add = clip_line_entity(entity, shp_polygon, msp)
                elif entity_type == 'LWPOLYLINE':
                    to_delete, to_add = clip_lwpolyline_entity(entity, shp_polygon, msp)
                elif entity_type == 'CIRCLE':
                    to_delete, to_add = clip_circle_entity(entity, shp_polygon, msp)
                elif entity_type == 'ARC':
                    to_delete, to_add = clip_arc_entity(entity, shp_polygon, msp)
                elif entity_type in ['TEXT', 'MTEXT','DIMENSION']:
                    to_delete, to_add = clip_text_entity(entity, shp_polygon, msp)
                elif entity_type == 'HATCH':
                    to_delete, to_add = clip_hatch_entity(entity, shp_polygon, msp)
                elif entity_type == 'POINT':
                    to_delete, to_add = clip_point_entity(entity, shp_polygon, msp)
                else:
                    to_delete, to_add = clip_text_entity(entity, shp_polygon, msp)
                all_to_delete.extend(to_delete)
                all_to_add.extend(to_add)
                
                # 更新进度
                if (idx + 1) % 1000 == 0 or idx == entity_count - 1:
                    progress = 15 + int((idx + 1) / entity_count * 70)
                    with task_lock:
                        task.progress = progress
                        task.message = f'已处理 {idx + 1}/{entity_count} 个实体...'
        
        # 4. 删除实体
        with task_lock:
            task.message = '正在删除实体...'
            task.progress = 85
        
        deleted_count = 0
        for entity in all_to_delete:
            try:
                msp.delete_entity(entity)
                deleted_count += 1
            except:
                pass
        for entity in all_to_add:
            try:
                msp.add_entity(entity)
            except:
                pass
        # 5. 保存文件
        with task_lock:
            task.message = '正在保存文件...'
            task.progress = 95
        
        # 确保输出目录存在
        output_dir = os.path.dirname(task.output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        doc.saveas(task.output_path)
        
        # 6. 完成
        with task_lock:
            task.status = 'completed'
            task.progress = 100
            task.message = f'裁剪完成! 删除{deleted_count}个实体, 添加{len(all_to_add)}个实体'
            task.end_time = datetime.now()
        
        logger.info(f"任务 {task_id} 完成")
        
    except Exception as e:
        logger.error(f"任务 {task_id} 失败: {e}")
        with task_lock:
            task.status = 'failed'
            task.error = str(e)
            task.message = f'裁剪失败: {str(e)}'
            task.end_time = datetime.now()


class ClipGui:
    def __init__(self):
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox

        self.tk = tk
        self.ttk = ttk
        self.filedialog = filedialog
        self.messagebox = messagebox

        self.root = tk.Tk()
        self.root.title("CAD裁剪工具")
        self.root.geometry("720x360")
        self.root.resizable(False, False)

        self.input_var = tk.StringVar()
        self.shp_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.status_var = tk.StringVar(value="请选择CAD和SHP文件")
        self.progress_var = tk.IntVar(value=0)

        self._task_id = None
        self._running = False

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        ttk = self.ttk
        main = ttk.Frame(self.root, padding=16)
        main.grid(row=0, column=0, sticky="nsew")

        ttk.Label(main, text="CAD文件:").grid(row=0, column=0, sticky="w")
        ttk.Entry(main, textvariable=self.input_var, width=70).grid(row=0, column=1, padx=8)
        self.cad_btn = ttk.Button(main, text="选择CAD", command=self._choose_cad)
        self.cad_btn.grid(row=0, column=2)

        ttk.Label(main, text="SHP文件:").grid(row=1, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(main, textvariable=self.shp_var, width=70).grid(row=1, column=1, padx=8, pady=(10, 0))
        self.shp_btn = ttk.Button(main, text="选择SHP", command=self._choose_shp)
        self.shp_btn.grid(row=1, column=2, pady=(10, 0))

        ttk.Label(main, text="输出文件:").grid(row=2, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(main, textvariable=self.output_var, width=70).grid(row=2, column=1, padx=8, pady=(10, 0))
        self.out_btn = ttk.Button(main, text="保存位置", command=self._choose_output)
        self.out_btn.grid(row=2, column=2, pady=(10, 0))

        self.start_btn = ttk.Button(main, text="开始裁剪", command=self._start_clip)
        self.start_btn.grid(row=3, column=1, pady=(16, 0), sticky="e")

        self.progress = ttk.Progressbar(main, maximum=100, variable=self.progress_var)
        self.progress.grid(row=4, column=0, columnspan=3, pady=(16, 4), sticky="ew")

        ttk.Label(main, textvariable=self.status_var).grid(row=5, column=0, columnspan=3, sticky="w")

    def _choose_cad(self):
        path = self.filedialog.askopenfilename(
            title="选择CAD文件",
            filetypes=[("CAD文件", "*.dxf;*.dwg"), ("DXF文件", "*.dxf"), ("DWG文件", "*.dwg"), ("全部文件", "*.*")]
        )
        if path:
            self.input_var.set(path)
            if not self.output_var.get():
                self.output_var.set(build_default_output_path(path))

    def _choose_shp(self):
        path = self.filedialog.askopenfilename(
            title="选择SHP文件",
            filetypes=[("SHP文件", "*.shp"), ("全部文件", "*.*")]
        )
        if path:
            self.shp_var.set(path)

    def _choose_output(self):
        input_path = self.input_var.get().strip()
        default_path = build_default_output_path(input_path) if input_path else ""
        ext = os.path.splitext(default_path)[1] if default_path else ".dxf"
        path = self.filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=ext,
            initialfile=os.path.basename(default_path) if default_path else "",
            filetypes=[("CAD文件", "*.dxf;*.dwg"), ("DXF文件", "*.dxf"), ("DWG文件", "*.dwg"), ("全部文件", "*.*")]
        )
        if path:
            self.output_var.set(path)

    def _start_clip(self):
        if self._running:
            return

        input_path = self.input_var.get().strip()
        shp_path = self.shp_var.get().strip()
        output_path = self.output_var.get().strip()

        if not input_path or not shp_path:
            self.messagebox.showwarning("提示", "请先选择CAD文件和SHP文件")
            return

        if not os.path.exists(input_path):
            self.messagebox.showerror("错误", f"CAD文件不存在: {input_path}")
            return

        if not os.path.exists(shp_path):
            self.messagebox.showerror("错误", f"SHP文件不存在: {shp_path}")
            return

        if not output_path:
            output_path = build_default_output_path(input_path)
            self.output_var.set(output_path)

        if os.path.abspath(output_path) == os.path.abspath(input_path):
            self.messagebox.showwarning("提示", "输出文件不能与输入文件相同")
            return

        task_id = str(uuid.uuid4())
        task = ClipTask(task_id, input_path, shp_path, output_path)

        with task_lock:
            tasks[task_id] = task

        thread = threading.Thread(target=process_clip_task, args=(task_id,))
        thread.daemon = True
        thread.start()

        self._task_id = task_id
        self._running = True
        self._set_running_state(True)
        self.status_var.set("任务已开始...")
        self.progress_var.set(0)
        self.root.after(200, self._poll_task)

    def _poll_task(self):
        if not self._running or not self._task_id:
            return

        with task_lock:
            task = tasks.get(self._task_id)

        if task:
            self.progress_var.set(task.progress)
            self.status_var.set(task.message)

            if task.status in ('completed', 'failed'):
                self._running = False
                self._set_running_state(False)

                if task.status == 'completed':
                    self.messagebox.showinfo("完成", f"裁剪完成，输出文件:\n{task.output_path}")
                else:
                    self.messagebox.showerror("失败", f"裁剪失败:\n{task.error}")

                with task_lock:
                    tasks.pop(self._task_id, None)
                return

        self.root.after(200, self._poll_task)

    def _set_running_state(self, running):
        state = "disabled" if running else "normal"
        self.cad_btn.config(state=state)
        self.shp_btn.config(state=state)
        self.out_btn.config(state=state)
        self.start_btn.config(state=state)

    def _on_close(self):
        if self._running:
            if not self.messagebox.askyesno("确认", "任务正在运行，确定要退出吗？"):
                return
        self.root.destroy()

    def run(self):
        self.root.mainloop()


@app.route('/api/clip', methods=['POST'])
def clip_cad():
    """
    裁剪接口
    POST /api/clip
    Body: {
        "input_dxf": "输入DXF文件路径",
        "shp_path": "SHP文件路径",
        "output_path": "输出DXF文件路径"
    }
    """
    try:
        data = request.get_json()
        
        # 验证参数
        input_dxf = data.get('input_dxf')
        shp_path = data.get('shp_path')
        output_path = data.get('output_path')
        
        if not all([input_dxf, shp_path, output_path]):
            return jsonify({
                'success': False,
                'message': '缺少必要参数: input_dxf, shp_path, output_path'
            }), 400
        
        # 验证文件是否存在
        if not os.path.exists(input_dxf):
            return jsonify({
                'success': False,
                'message': f'输入DXF文件不存在: {input_dxf}'
            }), 400
        
        if not os.path.exists(shp_path):
            return jsonify({
                'success': False,
                'message': f'SHP文件不存在: {shp_path}'
            }), 400
        
        # 创建任务
        task_id = str(uuid.uuid4())
        task = ClipTask(task_id, input_dxf, shp_path, output_path)
        
        with task_lock:
            tasks[task_id] = task
        
        # 启动异步任务
        thread = threading.Thread(target=process_clip_task, args=(task_id,))
        thread.daemon = True
        thread.start()
        
        logger.info(f"创建任务 {task_id}")
        
        return jsonify({
            'success': True,
            'message': '任务已创建',
            'task_id': task_id
        }), 200
        
    except Exception as e:
        logger.error(f"创建任务失败: {e}")
        return jsonify({
            'success': False,
            'message': f'创建任务失败: {str(e)}'
        }), 500


@app.route('/api/progress/<task_id>', methods=['GET'])
def get_progress(task_id):
    """
    查询进度接口
    GET /api/progress/<task_id>
    """
    try:
        with task_lock:
            task = tasks.get(task_id)
        
        if not task:
            return jsonify({
                'success': False,
                'message': '任务不存在'
            }), 404
        
        elapsed_time = None
        if task.start_time:
            end = task.end_time if task.end_time else datetime.now()
            elapsed_time = (end - task.start_time).total_seconds()
        
        return jsonify({
            'success': True,
            'task_id': task_id,
            'status': task.status,
            'progress': task.progress,
            'total': task.total,
            'message': task.message,
            'error': task.error,
            'elapsed_time': elapsed_time
        }), 200
        
    except Exception as e:
        logger.error(f"查询进度失败: {e}")
        return jsonify({
            'success': False,
            'message': f'查询进度失败: {str(e)}'
        }), 500


@app.route('/api/download/<task_id>', methods=['GET'])
def download_file(task_id):
    """
    下载裁剪后的文件接口
    GET /api/download/<task_id>
    """
    try:
        with task_lock:
            task = tasks.get(task_id)
        
        if not task:
            return jsonify({
                'success': False,
                'message': '任务不存在'
            }), 404
        
        if task.status != 'completed':
            return jsonify({
                'success': False,
                'message': f'任务未完成,当前状态: {task.status}'
            }), 400
        
        if not os.path.exists(task.output_path):
            return jsonify({
                'success': False,
                'message': '输出文件不存在'
            }), 404
        
        filename = os.path.basename(task.output_path)
        
        return send_file(
            task.output_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/octet-stream'
        )
        
    except Exception as e:
        logger.error(f"下载文件失败: {e}")
        return jsonify({
            'success': False,
            'message': f'下载文件失败: {str(e)}'
        }), 500

# 解析命令行参数
def parse_args():
    parser = argparse.ArgumentParser(description='CAD clipper with GUI or Flask server')
    parser.add_argument('--port', type=int, default=5000, help='Port number to run the Flask app on')
    parser.add_argument('--server', action='store_true', help='Run Flask server mode')
    parser.add_argument('--gui', action='store_true', help='Run desktop GUI mode')
    return parser.parse_args()

@app.route('/api/tasks', methods=['GET'])
def list_tasks():
    """
    列出所有任务
    GET /api/tasks
    """
    try:
        with task_lock:
            task_list = []
            for task_id, task in tasks.items():
                task_list.append({
                    'task_id': task_id,
                    'status': task.status,
                    'progress': task.progress,
                    'message': task.message,
                    'start_time': task.start_time.isoformat() if task.start_time else None,
                    'end_time': task.end_time.isoformat() if task.end_time else None
                })
        
        return jsonify({
            'success': True,
            'tasks': task_list
        }), 200
        
    except Exception as e:
        logger.error(f"列出任务失败: {e}")
        return jsonify({
            'success': False,
            'message': f'列出任务失败: {str(e)}'
        }), 500


@app.route('/api/health', methods=['GET'])
def health_check():
    """健康检查接口"""
    return jsonify({
        'success': True,
        'message': 'CAD裁剪服务运行正常',
        'timestamp': datetime.now().isoformat()
    }), 200


if __name__ == '__main__':
    args = parse_args()

    run_gui = args.gui or (getattr(sys, 'frozen', False) and not args.server)
    if run_gui:
        ClipGui().run()
    else:
        logger.info("="*60)
        logger.info("CAD裁剪服务启动 指定端口 --port 5000")
        logger.info("="*60)
        logger.info("API接口:")
        logger.info("  POST   /api/clip              - 创建裁剪任务")
        logger.info("  GET    /api/progress/<task_id> - 查询任务进度")
        logger.info("  GET    /api/download/<task_id> - 下载裁剪结果")
        logger.info("  GET    /api/tasks             - 列出所有任务")
        logger.info("  GET    /api/health            - 健康检查")
        logger.info("="*60)
        app.run(host='0.0.0.0', port=args.port, debug=False, threaded=True)
