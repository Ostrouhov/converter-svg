import math
import numpy as np
import svgwrite

xpath = './/{http://www.w3.org/2000/svg}'


class StartCoordinates:
    def __init__(self, *, x_start: str, y_start: str):
        self.x_start = x_start
        self.y_start = y_start

class Figure:
    def __init__(self, *, x_start: str, y_start: str, color: str = 'gray'):
        self.start_coordinates = StartCoordinates(x_start=x_start, y_start=y_start)
        self.color = color

class Point(Figure):
    def __int__(self, *, x_start: str, y_start: str, color: str):
        super().__init__(x_start=x_start, y_start=y_start, color=color)

    def draw(self, dwg) -> None:
        dwg.add(dwg.circle(center=(self.start_coordinates.x_start, self.start_coordinates.y_start), r=2, fill=self.color))

class Rectangle(Figure):
    def __init__(self, *, x_start: str, y_start: str, color: str, width: str, height: str) -> None:
        super().__init__(x_start=x_start, y_start=y_start, color=color)
        self.width = width
        self.height = height

    @staticmethod
    def find_values_and_create_instances(root):

        values = list()
        # Ищем все теги <rect> и проверяем атрибут id
        for rect in root.findall(f'{xpath}rect'):
            rect_id = rect.get('id')
            if rect_id and rect_id.startswith('Rect'):
                r = Rectangle(x_start=rect.get('x'), y_start=rect.get('y'), color=rect.get('fill'),
                              width=rect.get('width'), height=rect.get('height'))
                values.append(r)

        return values

    @staticmethod
    def stylize(all_rect):
        style_rect = list()

        for r in all_rect:
            if r.color == "#c0c0c0":
                r.color = "#616161"
                style_rect.append(r)
            elif r.color == "#a4c4ff":
                r.color = "#3E5F8A"
                style_rect.append(r)

        return style_rect

    def draw(self, dwg):
        dwg.add(dwg.rect(insert=(self.start_coordinates.x_start, self.start_coordinates.y_start),
                         size=(self.width, self.height), fill=self.color, stroke='black', stroke_width=1))

class Polygon():
    def __init__(self, *, points: list, color: str):
        self.points = points
        self.color = color

    @staticmethod
    def find_values_and_create_instances(root):

        values = list()
        for pol in root.findall(f'{xpath}polygon'):
            pol_id = pol.get('id')
            if pol_id and pol_id.startswith('Polygon'):
                points = pol.get('points').split(" ")
                points_list = list()
                for pn in points:
                    points_list.append(pn.split(","))
                values.append(Polygon(points=points_list, color=pol.get('fill')))

        return values

    def draw(self, dwg):
        dwg.add(dwg.polygon(points=self.points, fill=self.color))

class Text(Figure):
    def __init__(self, *, x_start: str, y_start: str, font_size: str, content: str, color: str = 'gray') -> None:
        super().__init__(x_start=x_start, y_start=y_start, color=color)
        self.font_size = font_size
        self.content = content

    @staticmethod
    def find_values_and_create_instances(root):
        values = list()

        for t in root.findall(f'{xpath}text'):
            t_id = t.get('id')
            if t_id and t_id.startswith('Text'):
                text = Text(x_start=t.get('x'), y_start=t.get('y'), font_size=t.get('font-size'), content=t.text)
                values.append(text)

        for ts in root.findall(f'{xpath}tspan'):
            tspan = Text(x_start=ts.get('x'), y_start=ts.get('y'), font_size=ts.get('font-size'), content=ts.text)
            values.append(tspan)

        return values

    def draw(self, dwg) -> None:
        txt = dwg.text(self.content, insert=(self.start_coordinates.x_start, self.start_coordinates.y_start), fill=self.color)
        txt['text-anchor'] = 'middle'
        dwg.add(txt)

class Valve(Figure):
    def __int__(self, *, x: str, y: str, color: str):
        super().__init__(x_start=x, y_start=y, color=color)

    @staticmethod
    def find_values_and_create_instances(root):
        values = list()

        for v in root.findall(f'{xpath}g'):
            v_id = v.get('id')
            if v_id and v_id.startswith('ZD_CS_U'):
                vl = Valve(x_start=v.get('x'), y_start=v.get('y'))
                values.append(vl)

        return values

    def align(self, line_y, line_x1, line_x2):
        if ((int(self.start_coordinates.y_start) <= int(line_y) <= int(self.start_coordinates.y_start) + 48) and
                (int(line_x1) <= int(self.start_coordinates.x_start) <= int(line_x2))):
            self.start_coordinates.y_start = int(line_y) - 25
    def draw(self, dwg):
        dwg.add(dwg.path(d="M16 24.5L0.25 33.5933L0.25 15.4067L16 24.5Z "
                           "M16 24.5L31.75 15.4067V33.5933L16 24.5Z "
                           "M26.5 10.5C26.5 4.70101 21.799 0 16 0C10.201 0 5.5 4.70101 5.5 10.5L26.5 10.5Z "
                           "M17 0.047V0H16C16.3373 0 16.6709 0.015903 17 0.047ZM17 10.5H16V24.5H17V10.5Z", fill=self.color, stroke="black",
                         transform=f"translate({self.start_coordinates.x_start},{self.start_coordinates.y_start})"))

'''
Первый вариант задвижек
"M23 35 L0 47 V22 L23 35 Z"
"M23 35 L45 22 V47 L23 35 Z"
"M38 15 C38 6 31 0 23 0 C14 0 8 6 8 15 L38 15 Z"
"M24 0 V0 H23 C23 0 23 0 24 0 Z M24 15 H23 V35 H24 V15 Z"
'''


class Line(Figure):
    def __init__(self, *, x_start: str, y_start: str, x_end: str, y_end: str,
                 color: str, stroke_width: str, stroke_dasharray: str):
        super().__init__(x_start=x_start, y_start=y_start, color=color)
        self.x_end = x_end
        self.y_end = y_end
        self.stroke_width = stroke_width
        self.stroke_dasharray = stroke_dasharray

    @staticmethod
    def find_values_and_create_instances(root):
        values = list()

        for ln in root.findall(f'{xpath}line'):
            line = Line(x_start=ln.get('x1'), y_start=ln.get('y1'), x_end=ln.get('x2'), y_end=ln.get('y2'),
                        color=ln.get('stroke'), stroke_width=ln.get('stroke-width'), stroke_dasharray=ln.get('stroke-dasharray'))
            values.append(line)

        return values

    @staticmethod
    def stylize(all_lines):
        lines1, lines2, lines3, lines4, lines5 = list(), list(), list(), list(), list()
        for l in all_lines:
            if l.color == "#ffff00":
                l.color = "#ada110"
                l.stroke_width = 2
                lines1.append(l)
            elif l.color == "#800080":
                l.color = "#1a001a"
                l.stroke_width = 2
                lines2.append(l)
            elif l.color == "#00ffff":
                l.color = "#006666"
                l.stroke_width = 2
                lines3.append(l)
            elif l.color == "#008000":
                l.stroke_width = 2
                lines4.append(l)
            elif l.stroke_dasharray is not None:
                l.stroke_dasharray = "10 5"
                lines5.append(l)

        return [lines1, lines2, lines3, lines4, lines5]

    def find_near_lines(self, lines):
        near_points = list()
        min_d = 9999
        x = 0
        y = 0
        x_out = 0
        y_out = 0

        for l in lines:
            if self.start_coordinates.x_start != l.start_coordinates.x_start and self.start_coordinates.y_start != l.start_coordinates.y_start:
                d = math.sqrt(math.pow(int(l.start_coordinates.x_start) - int(self.start_coordinates.x_start), 2)
                              + math.pow(int(l.start_coordinates.y_start) - int(self.start_coordinates.y_start), 2))
                if d < min_d:
                    min_d = d
                    x = l.start_coordinates.x_start
                    y = l.start_coordinates.y_start
                    x_out = l.x_end
                    y_out = l.y_end
                d = math.sqrt(math.pow(int(l.x_end) - int(self.start_coordinates.x_start), 2)
                              + math.pow(int(l.y_end) - int(self.start_coordinates.y_start), 2))
                if d < min_d:
                    min_d = d
                    x = l.x_end
                    y = l.y_end
                    x_out = l.start_coordinates.x_start
                    y_out = l.start_coordinates.y_start

        point1 = Point(x_start=self.start_coordinates.x_start, y_start=self.start_coordinates.y_start) # Начальная точка (НТ)
        point2 = Point(x_start=x, y_start=y) # Самая близка к НТ точка (БТ)
        point3 = Point(x_start=x_out, y_start=y_out) # Конечная точка для БТ, для создания линии
        near_points.append(point1)
        near_points.append(point2)
        near_points.append(point3)
        # near_points[f'{point1.x},{point1.y}'] = f'{point2.x},{point2.y}'

        min_d = 9999
        x = 0
        y = 0
        x_out = 0
        y_out = 0
        # print(f'({point1.x}, {point1.y})---({point2.x}, {point2.y})')

        for l in lines:
            if self.x_end != l.x_end and self.y_end != l.y_end:
                d = math.sqrt(math.pow(int(l.start_coordinates.x_start) - int(self.x_end), 2)
                              + math.pow(int(l.start_coordinates.y_start) - int(self.y_end), 2))
                if d < min_d:
                    min_d = d
                    x = l.start_coordinates.x_start
                    y = l.start_coordinates.y_start
                    x_out = l.x_end
                    y_out = l.y_end
                d = math.sqrt(math.pow(int(l.x_end) - int(self.x_end), 2)
                              + math.pow(int(l.y_end) - int(self.y_end), 2))
                if d < min_d:
                    min_d = d
                    x = l.x_end
                    y = l.y_end
                    x_out = l.start_coordinates.x_start
                    y_out = l.start_coordinates.y_start

        point1 = Point(x_start=self.x_end, y_start=self.y_end)
        point2 = Point(x_start=x, y_start=y)
        point3 = Point(x_start=x_out, y_start=y_out)
        near_points.append(point1)
        near_points.append(point2)
        near_points.append(point3)

        return near_points

    def align(self, lines):
        near_line = Line.find_near_lines(self, lines)

        intersection_point = find_intersection(x1=int(near_line[0].start_coordinates.x_start), y1=int(near_line[0].start_coordinates.y_start),
                                               x2=int(near_line[3].start_coordinates.x_start), y2=int(near_line[3].start_coordinates.y_start),
                                               x3=int(near_line[1].start_coordinates.x_start), y3=int(near_line[1].start_coordinates.y_start),
                                               x4=int(near_line[2].start_coordinates.x_start), y4=int(near_line[2].start_coordinates.y_start))
        if (intersection_point is not None) and (0 < intersection_point[0] < 2000 and 0 < intersection_point[1] < 2000):
            if (abs(int(self.start_coordinates.x_start) - intersection_point[0]) < 10) and (abs(int(self.start_coordinates.y_start) - intersection_point[1]) < 10):
                self.start_coordinates.x_start = intersection_point[0]
                self.start_coordinates.y_start = intersection_point[1]

        intersection_point = find_intersection(x1=int(near_line[0].start_coordinates.x_start), y1=int(near_line[0].start_coordinates.y_start),
                                               x2=int(near_line[3].start_coordinates.x_start), y2=int(near_line[3].start_coordinates.y_start),
                                               x3=int(near_line[4].start_coordinates.x_start), y3=int(near_line[4].start_coordinates.y_start),
                                               x4=int(near_line[5].start_coordinates.x_start), y4=int(near_line[5].start_coordinates.y_start))

        if (intersection_point is not None) and (0 < intersection_point[0] < 2000 and 0 < intersection_point[1] < 2000):
            if (abs(int(self.x_end) - intersection_point[0]) < 10) and (abs(int(self.y_end) - intersection_point[1]) < 10):
                self.x_end = intersection_point[0]
                self.y_end = intersection_point[1]


        # for l in lines:
        #     min_y = min(intself.start_coordinates.y_start - l.y_end)
        #
        # for l in lines:
        #     # if int(self.start_coordinates.y_start) == int(self.y_end):
        #     #     if int(self.start_coordinates.x_start) > int(l.x_end) > int(self.start_coordinates.x_start) - 10:
        #     #         self.start_coordinates.x_start = l.x_end
        #
        #     if int(self.start_coordinates.x_start) == int(self.x_end):
        #         if int(self.start_coordinates.y_start) < int(l.y_end) < int(self.start_coordinates.y_start) + 10 and (min_y == l.y_end):
        #             self.start_coordinates.y_start = l.y_end

    def draw(self, dwg):
        if self.stroke_dasharray is None:
            self.stroke_dasharray = "0 0"
        dwg.add(dwg.line(start=(self.start_coordinates.x_start, self.start_coordinates.y_start), end=(self.x_end, self.y_end),
                         stroke=self.color, stroke_width=self.stroke_width, stroke_dasharray=self.stroke_dasharray))

def find_intersection(*, x1: int, y1: int, x2: int, y2: int, x3: int, y3: int, x4: int, y4: int):
    # Переводим координаты начальных и конечных точек прямых в массивы numpy
    line1_start = np.array([x1, y1])
    line1_end = np.array([x2, y2])
    line2_start = np.array([x3, y3])
    line2_end = np.array([x4, y4])

    # Вычисляем векторы, представляющие прямые
    v1 = line1_end - line1_start
    v2 = line2_end - line2_start

    # Создаем матрицу коэффициентов для системы линейных уравнений
    matrix = np.vstack((v1, v2)).T
    constants = line2_start - line1_start

    try:
        # Решаем систему линейных уравнений
        t1, t2 = np.linalg.solve(matrix, constants)

        # Находим координаты точки пересечения
        intersection_point = line1_start + t1 * v1

        return intersection_point[0], intersection_point[1]

    except np.linalg.LinAlgError:
        # Обработка ситуации, когда матрица является сингулярной
        return None