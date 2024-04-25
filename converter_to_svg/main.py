import lxml.etree as et
import svgwrite
import objects as ob
import tkinter as tk
from tkinter import filedialog

def find_all_elements(*, cl, root):
    return cl.find_values_and_create_instances(root)

def find_style_elements(*, cl, list_of_all_elements):
    return cl.stylize(list_of_all_elements)

def draw_elements(element, dwg):
    element.draw(dwg)

def load_files():
    filenames = filedialog.askopenfilenames(title="Выберите файлы",
                                            filetypes=(("Svg files", "*.svg"), ("All files", "*.*")))
    if filenames:
        files_path = list(filenames)
        for file in files_path:
            try:
                tree = et.parse(file)
                root = tree.getroot()
            except OSError:
                print(f"File {file} not found.")
            except et.ParseError:
                print(f"Unable to parse {file} as an SVG file.")

            dwg = svgwrite.Drawing(f'{file.split("/")[-1][:-4]}_OUTPUT.svg', profile='tiny')

            all_rectangles = find_all_elements(cl=ob.Rectangle, root=root)
            rectangles = find_style_elements(cl=ob.Rectangle, list_of_all_elements=all_rectangles)
            all_lines = find_all_elements(cl=ob.Line, root=root)
            lines_lists = find_style_elements(cl=ob.Line, list_of_all_elements=all_lines)
            text = find_all_elements(cl=ob.Text, root=root)
            valve = find_all_elements(cl=ob.Valve, root=root)
            all_polygons = find_all_elements(cl=ob.Polygon, root=root)

            for r in all_rectangles:
                draw_elements(r, dwg)

            for polyg in all_polygons:
                draw_elements(polyg, dwg)

            # for l in all_lines:
            #     draw_elements(l)

            ln = list()

            for i in range(len(lines_lists)):
                for l in lines_lists[i]:
                    ob.Line.align(l, lines_lists[i])
                    ln.append(l)

            for l in ln:
                draw_elements(l, dwg)

            for t in text:
                draw_elements(t, dwg)

            for v in valve:
                for l in ln:
                    ob.Valve.align(v, l.start_coordinates.y_start, l.start_coordinates.x_start, l.x_end)
                draw_elements(v, dwg)

            dwg.save()

        print("Выбранные файлы:", files_path)


if __name__ == "__main__":

    win = tk.Tk()
    win.title('Менеджер открытия svg-файлов')
    win.geometry("400x300")

    win.btn_load = tk.Button(win, text="Загрузить svg-файлы", command=load_files)
    win.btn_load.pack(pady=10)

    win.mainloop()




    # ln = list()
    #
    # for l in lines:
    #     if l.color == "#ada110" or l.color == "#800080":
    #         ln.append(l)
    # for l in ln:
    #     Line.align(l, ln)
    # for l in lines:
    #     draw_elements(l)
    #
    # for t in text:
    #     draw_elements(t)
    #
    # for v in valve:
    #     for l in lines:
    #         Valve.align(v, l.start_coordinates.y_start, l.start_coordinates.x_start, l.x_end)
    #     draw_elements(v)
