import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import pandas as pd
import numpy as np
from abc import ABC, abstractmethod

def parse_excel_for_chart(file_content):
    """解析Excel文件为图表数据"""
    try:
        # 读取所有sheet
        excel_file = pd.ExcelFile(BytesIO(file_content))
        sheet_names = excel_file.sheet_names
        
        # 返回所有sheet的数据
        sheets_data = {}
        for sheet in sheet_names:
            df = pd.read_excel(BytesIO(file_content), sheet_name=sheet)
            sheets_data[sheet] = {
                'columns': df.columns.tolist(),
                'data': df.to_dict('records')
            }
        
        return {
            'sheets': sheet_names,
            'data': sheets_data
        }
        
    except Exception as e:
        raise Exception(f"解析Excel文件失败: {str(e)}")

class ChartBase(ABC):
    """图表基类"""
    def __init__(self, df, columns, chart_title=None, show_labels=True, axis_settings=None):
        self.df = df
        self.columns = columns
        self.chart_title = chart_title
        self.show_labels = show_labels
        self.axis_settings = axis_settings or {}
        
    @abstractmethod
    def validate(self):
        """验证数据"""
        pass
        
    @abstractmethod
    def create(self):
        """创建图表"""
        pass
        
    def get_layout_updates(self):
        """获取基础布局更新"""
        layout = {
            'showlegend': True,
            'template': 'plotly_white',
            'margin': dict(t=80, l=80, r=50, b=80),
            'dragmode': False,
            'annotations': [],
            'showlegend': True,
            'modebar': {
                'remove': ['editInChartStudio', 'editable']
            },
            'title': {
                'text': None
            }
        }
        
        # 处理图表标题
        if self.chart_title:
            title_settings = self.axis_settings.get('titleSettings', {})
            font_size = int(title_settings.get('fontSize', 24))
            
            layout['title'].update({
                'text': self.chart_title,
                'font': {
                    'size': font_size,
                    'color': title_settings.get('color', '#000000')
                },
                'x': 0.5,
                'y': 0.95,
                'xanchor': 'center',
                'yanchor': 'top',
                'pad': {'t': 20},
                'automargin': True
            })
        
        return layout

    def _get_axis_range(self, min_val, max_val, data_min, data_max):
        """获取坐标轴范围"""
        range_min = None
        range_max = None
        
        # 如果设置了最小值
        if min_val not in [None, '']:
            try:
                range_min = float(min_val)
            except ValueError:
                pass
        
        # 如果设置了最大值
        if max_val not in [None, '']:
            try:
                range_max = float(max_val)
            except ValueError:
                pass
        
        # 如果都没有设置，返回None
        if range_min is None and range_max is None:
            return None
            
        # 如果只设置了最小值
        if range_min is not None and range_max is None:
            range_max = data_max * 1.1  # 默认最大值留10%空间
            
        # 如果只设置了最大值
        if range_min is None and range_max is not None:
            range_min = data_min * 0.9  # 默认最小值留10%空间
            
        return [range_min, range_max]

    def _get_axis_settings(self, axis_type, settings):
        """获取坐标轴设置"""
        title_settings = settings.get(f'{axis_type}TitleSettings', {})
        
        # 确保字体大小是整数且在合理范围内
        font_size = title_settings.get('fontSize')
        try:
            font_size = int(font_size) if font_size else 14
            font_size = max(10, min(font_size, 32))
        except (ValueError, TypeError):
            font_size = 14
        
        # 基础设置，包含轴标题设置
        axis_settings = {
            'title': {
                'text': settings.get(f'{axis_type}AxisTitle', ''),
                'font': {
                    'size': font_size,
                    'color': title_settings.get('color', '#000000')
                },
                'standoff': 25
            },
            'tickfont': {'size': 12},
            'showgrid': True,
            'gridwidth': 1,
            'gridcolor': '#E1E1E1',
            'automargin': True
        }
        
        # 使用 annotation 作为轴标题的补充
        title_annotation = {
            'text': settings.get(f'{axis_type}AxisTitle', ''),
            'font': {
                'size': font_size,  # 使用相同的字体大小
                'color': title_settings.get('color', '#000000')
            },
            'showarrow': False,
            'captureevents': True
        }
        
        if axis_type == 'x':
            title_annotation.update({
                'x': 0.5,
                'y': -0.15,
                'xref': 'paper',
                'yref': 'paper',
                'xanchor': 'center',
                'yanchor': 'top'
            })
            # X轴特有设置
            axis_settings['title'].update({
                'standoff': 40,
                'automargin': True
            })
        else:
            title_annotation.update({
                'x': -0.15,
                'y': 0.5,
                'xref': 'paper',
                'yref': 'paper',
                'xanchor': 'right',
                'yanchor': 'middle',
                'textangle': -90
            })
            # Y轴特有设置
            axis_settings['title'].update({
                'standoff': 40,
                'automargin': True
            })
        
        # 将标题添加到布局的 annotations 中
        if not hasattr(self, '_annotations'):
            self._annotations = []
        self._annotations.append(title_annotation)
        
        return axis_settings

    def create(self):
        """创建图表的基础方法"""
        self.validate()
        fig = self._create_figure()
        
        # 更新布局
        layout = self.get_layout_updates()
        
        # 合并所有 annotations
        if hasattr(self, '_annotations'):
            layout['annotations'] = layout.get('annotations', []) + self._annotations
        
        # 禁用编辑模式
        fig.update_layout(
            **layout,
            editrevision=False,  # 禁用编辑修订
            editable=False,      # 禁用编辑模式
            showEditButtons=False  # 隐藏编辑按钮
        )
        
        return fig

class BarChart(ChartBase):
    """柱状图"""
    def __init__(self, df, columns, chart_title=None, show_labels=True, axis_settings=None):
        super().__init__(df, columns, chart_title, show_labels)
        self.axis_settings = axis_settings or {}
        
    def validate(self):
        """验证数据"""
        if len(self.columns) < 2:
            raise ValueError("柱状图需要至少选择两列数据（X轴和Y轴）")
            
        # 验证Y轴数据是否为数值类型
        for col in self.columns[1:]:
            if not pd.api.types.is_numeric_dtype(self.df[col]):
                raise ValueError(f"Y轴列 '{col}' 必须是数值类型")
                
    def create(self):
        """创建柱状图"""
        self.validate()
        
        fig = px.bar(
            self.df,
            x=self.columns[0],
            y=self.columns[1:],
            barmode='group',
            text_auto='.2f',
            labels={
                self.columns[0]: self.axis_settings.get('xAxisTitle', self.columns[0]),
                'value': self.axis_settings.get('yAxisTitle', '数值'),
                'variable': '数据系列'
            }
        )
        
        # 更新布局
        layout = self.get_layout_updates()
        
        # 获取X轴和Y轴设置
        xaxis_settings = self._get_axis_settings('x', self.axis_settings)
        yaxis_settings = self._get_axis_settings('y', self.axis_settings)
        
        # 如果X轴是数值类型，设置范围
        if pd.api.types.is_numeric_dtype(self.df[self.columns[0]]):
            x_data_min = self.df[self.columns[0]].min()
            x_data_max = self.df[self.columns[0]].max()
            
            x_range = self._get_axis_range(
                self.axis_settings.get('xAxisMin'),
                self.axis_settings.get('xAxisMax'),
                x_data_min,
                x_data_max
            )
            if x_range:
                xaxis_settings['range'] = x_range
        
        # 设置Y轴范围
        y_data_min = min(self.df[col].min() for col in self.columns[1:])
        y_data_max = max(self.df[col].max() for col in self.columns[1:])
        
        y_range = self._get_axis_range(
            self.axis_settings.get('yAxisMin'),
            self.axis_settings.get('yAxisMax'),
            y_data_min,
            y_data_max
        )
        if y_range:
            yaxis_settings['range'] = y_range
        
        layout.update({
            'bargap': 0.15,
            'bargroupgap': 0.1,
            'xaxis': xaxis_settings,
            'yaxis': yaxis_settings
        })
        
        # 数据标签设置
        if self.show_labels:
            fig.update_traces(
                textposition='outside',
                texttemplate='%{y:.1f}',
                textfont=dict(
                    size=11,
                    color='black'
                ),
                textangle=0,
                cliponaxis=False
            )
        else:
            fig.update_traces(
                textposition='none',
                text=None
            )
        
        fig.update_layout(**layout)
        return fig

class LineChart(ChartBase):
    """折线图"""
    def __init__(self, df, columns, chart_title=None, show_labels=True, axis_settings=None):
        super().__init__(df, columns, chart_title, show_labels)
        self.axis_settings = axis_settings or {}
        
    def validate(self):
        """验证数据"""
        if len(self.columns) < 2:
            raise ValueError("折线图需要至少选择两列数据（X轴和Y轴）")
            
        # 验证Y轴数据是否为数值类型
        for col in self.columns[1:]:
            if not pd.api.types.is_numeric_dtype(self.df[col]):
                raise ValueError(f"Y轴列 '{col}' 必须是数值类型")
                
    def create(self):
        """创建折线图"""
        self.validate()
        
        fig = px.line(
            self.df,
            x=self.columns[0],
            y=self.columns[1:],
            markers=True,
            labels={
                self.columns[0]: self.axis_settings.get('xAxisTitle', self.columns[0]),
                'value': self.axis_settings.get('yAxisTitle', '数值'),
                'variable': '数据系列'
            }
        )
        
        # 更新布局
        layout = self.get_layout_updates()
        
        # 设置Y轴范围和标题
        yaxis_settings = {
            'title': self.axis_settings.get('yAxisTitle', '数值'),
            'zeroline': True,
            'gridwidth': 1,
            'gridcolor': '#E1E1E1'
        }
        
        # 获取Y轴数据范围
        y_data_min = min(self.df[col].min() for col in self.columns[1:])
        y_data_max = max(self.df[col].max() for col in self.columns[1:])
        
        # 设置Y轴范围
        y_range = self._get_axis_range(
            self.axis_settings.get('yAxisMin'),
            self.axis_settings.get('yAxisMax'),
            y_data_min,
            y_data_max
        )
        if y_range:
            yaxis_settings['range'] = y_range
            
        # 设置X轴范围和标题
        xaxis_settings = {
            'title': self.axis_settings.get('xAxisTitle', self.columns[0]),
            'tickangle': -45,
            'tickmode': 'auto',
            'nticks': 20
        }
        
        # 如果X轴是数值类型，设置范围
        if pd.api.types.is_numeric_dtype(self.df[self.columns[0]]):
            x_data_min = self.df[self.columns[0]].min()
            x_data_max = self.df[self.columns[0]].max()
            
            x_range = self._get_axis_range(
                self.axis_settings.get('xAxisMin'),
                self.axis_settings.get('xAxisMax'),
                x_data_min,
                x_data_max
            )
            if x_range:
                xaxis_settings['range'] = x_range
        
        layout.update({
            'xaxis': xaxis_settings,
            'yaxis': yaxis_settings
        })
        
        # 根据设置显示或隐藏数据标签
        if self.show_labels:
            fig.update_traces(
                textposition='top center',
                texttemplate='%{y:.1f}',
                textfont=dict(
                    size=11,
                    color='black'
                )
            )
        else:
            fig.update_traces(
                textposition='none',
                text=None
            )
        
        fig.update_layout(**layout)
        return fig

class PieChart(ChartBase):
    """饼图"""
    def __init__(self, df, columns, chart_title=None, show_labels=True, pie_settings=None):
        super().__init__(df, columns, chart_title, show_labels)
        self.pie_settings = pie_settings or {}
        
    def validate(self):
        if len(self.columns) != 2:
            raise ValueError("饼图需要选择两列数据（类别和数值）")
            
        if not pd.api.types.is_numeric_dtype(self.df[self.columns[1]]):
            raise ValueError(f"'{self.columns[1]}' 列必须是数值类型")
            
        if (self.df[self.columns[1]] < 0).any():
            raise ValueError("饼图不支持负数值")
            
        if (self.df[self.columns[1]] == 0).all():
            raise ValueError("数值列不能全为0")
            
    def create(self):
        self.validate()
        
        hole_size = self.pie_settings.get('holeSize', 0.3)
        show_percentage = self.pie_settings.get('showPercentage', True)
        show_values = self.pie_settings.get('showValues', True)
        show_labels = self.pie_settings.get('showLabels', True)
        
        fig = px.pie(
            self.df,
            names=self.columns[0],
            values=self.columns[1],
            hole=hole_size
        )
        
        # 配置标签显示
        text_info = []
        if show_labels:
            text_info.append('label')
        if show_percentage:
            text_info.append('percent')
        if show_values:
            text_info.append('value')
        
        if text_info:
            fig.update_traces(
                textposition='inside',
                textinfo='+'.join(text_info),
                textfont_size=12,
                insidetextorientation='radial'
            )
        else:
            fig.update_traces(
                textposition='none',
                textinfo='none'
            )
            
        layout = self.get_layout_updates()
        fig.update_layout(**layout)
        return fig

def create_chart(data, chart_type, columns, chart_title=None, show_labels=True, pie_settings=None, axis_settings=None):
    """创建图表的工厂函数"""
    try:
        df = pd.DataFrame(data)
        
        # 验证列是否存在
        for col in columns:
            if col not in df.columns:
                raise ValueError(f"列 '{col}' 不存在")
        
        # 根据图表类型创建相应的图表对象
        chart_classes = {
            'bar': BarChart,
            'line': LineChart,
            'pie': PieChart
        }
        
        chart_class = chart_classes.get(chart_type)
        if not chart_class:
            raise ValueError(f"不支持的图表类型: {chart_type}")
            
        # 创建图表对象
        if chart_type == 'pie':
            chart = chart_class(
                df=df,
                columns=columns,
                chart_title=chart_title,
                show_labels=show_labels,
                pie_settings=pie_settings
            )
        else:
            chart = chart_class(
                df=df,
                columns=columns,
                chart_title=chart_title,
                show_labels=show_labels,
                axis_settings=axis_settings
            )
            
        # 创建图表
        fig = chart.create()
        
        return fig.to_json()
        
    except Exception as e:
        raise Exception(f"创建图表失败: {str(e)}")

def export_chart(data, chart_type, columns, format='png'):
    """导出图表为图片"""
    try:
        df = pd.DataFrame(data)
        
        # 创建图表
        chart_result = create_chart(data, chart_type, columns)
        fig = go.Figure(chart_result)
        
        # 创建内存文件对象
        img_bytes = BytesIO()
        
        # 将图表保存为图片
        if format == 'png':
            fig.write_image(img_bytes, format='png')
        elif format == 'svg':
            fig.write_image(img_bytes, format='svg')
        else:
            raise ValueError(f"不支持的导出格式: {format}")
        
        img_bytes.seek(0)
        return img_bytes
        
    except Exception as e:
        raise Exception(f"导出图表失败: {str(e)}")

# 确保导出所有需要的函数
__all__ = ['parse_excel_for_chart', 'create_chart', 'export_chart']