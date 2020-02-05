CREATE OR REPLACE PACKAGE OD.as_xlsx
is
/**********************************************
 Пакет для работы с Excel файлами и выгрузки их на сервер
 Автор Мошков Р.А.
 Версия от 12/10/2019
******************************************************************************
******************************************** */
-- список констант для XLSX
-- виды окантовок ячейки
xl_b_none					constant varchar2 (128) := 'none';
xl_b_thin					constant varchar2 (128) := 'thin';
xl_b_medium					constant varchar2 (128) := 'medium';
xl_b_dashedconst			constant varchar2 (128) := 'dashed';
xl_b_dottedconst			constant varchar2 (128) := 'dotted';
xl_b_thick					constant varchar2 (128) := 'thick';
xl_b_double					constant varchar2 (128) := 'double';
xl_b_hair					constant varchar2 (128) := 'hair';
xl_b_mediumDashed			constant varchar2 (128) := 'mediumDashed';
xl_b_dashDot				constant varchar2 (128) := 'dashDot';
xl_b_mediumDashDot			constant varchar2 (128) := 'mediumDashDot';
xl_b_dashDotDot				constant varchar2 (128) := 'dashDotDot';
xl_b_mediumDashDotDot		constant varchar2 (128) := 'mediumDashDotDot';
xl_b_slantDashDot			constant varchar2 (128) := 'slantDashDot';
-- выравнивание по горизонтали
xl_ah_horizontal			constant varchar2 (128) := 'horizontal';
xl_ah_center				constant varchar2 (128) := 'center';
xl_ah_centerContinuous		constant varchar2 (128) := 'centerContinuous';
xl_ah_distributed			constant varchar2 (128) := 'distributed';
xl_ah_fill					constant varchar2 (128) := 'fill';
xl_ah_general				constant varchar2 (128) := 'general';
xl_ah_justify				constant varchar2 (128) := 'justify';
xl_ah_left					constant varchar2 (128) := 'left';
xl_ah_right					constant varchar2 (128) := 'right';
-- выравнивание по вертикали
xl_av_vertical				constant varchar2 (128) := 'vertical';
xl_av_bottom				constant varchar2 (128) := 'bottom';
xl_av_center				constant varchar2 (128) := 'center';
xl_av_distributed			constant varchar2 (128) := 'distributed';
xl_av_justify				constant varchar2 (128) := 'justify';
xl_av_top					constant varchar2 (128) := 'top';
-- типы данных ячейки
xl_tp_bool					constant varchar2 (128) := 'b';			-- логику
xl_tp_date					constant varchar2 (128) := 'd';			-- содержит дату в формате ISO 8601.
xl_tp_error					constant varchar2 (128) := 'e';			-- метка ошибки содержания.
xl_tp_inlineStr				constant varchar2 (128) := 'inlineStr';	-- содежит многострочный текст он не содержится в SharedString.
xl_tp_number				constant varchar2 (128) := 'n';			-- содержит число.
xl_tp_SharedString			constant varchar2 (128) := 's';			-- содежит однострочный текст он содержится в SharedString.
xl_tp_formula				constant varchar2 (128) := 'str';		-- формула
--
xl_inf_stop					constant varchar2 (128) := 'stop';
xl_inf_warning				constant varchar2 (128) := 'warning';
xl_inf_information			constant varchar2 (128) := 'information';
-- Стили заливки
xl_fs_darkDown				constant varchar2 (128) := 'darkDown';
xl_fs_darkGray				constant varchar2 (128) := 'darkGray';
xl_fs_darkGrid				constant varchar2 (128) := 'darkGrid';
xl_fs_darkHorizontal		constant varchar2 (128) := 'darkHorizontal';
xl_fs_darkTrellis			constant varchar2 (128) := 'darkTrellis';
xl_fs_darkUp				constant varchar2 (128) := 'darkUp';
xl_fs_darkVertical			constant varchar2 (128) := 'darkVertical';
xl_fs_gray0625				constant varchar2 (128) := 'gray0625';
xl_fs_gray125				constant varchar2 (128) := 'gray125';
xl_fs_lightDown				constant varchar2 (128) := 'lightDown';
xl_fs_lightGray				constant varchar2 (128) := 'lightGray';
xl_fs_lightGrid				constant varchar2 (128) := 'lightGrid';
xl_fs_lightHorizontal		constant varchar2 (128) := 'lightHorizontal';
xl_fs_lightTrellis			constant varchar2 (128) := 'lightTrellis';
xl_fs_lightUp				constant varchar2 (128) := 'lightUp';
xl_fs_lightVertical			constant varchar2 (128) := 'lightVertical';
xl_fs_mediumGray			constant varchar2 (128) := 'mediumGray';
xl_fs_none					constant varchar2 (128) := 'none';
xl_fs_solid					constant varchar2 (128) := 'solid';
-- Вертикальное выравнивание
xl_va_baseline				constant varchar2 (128) := 'baseline';
xl_va_superscript			constant varchar2 (128) := 'superscript';
xl_va_subscript				constant varchar2 (128) := 'subscript';
-- Шаблоны для колонтитулов
xl_hd_PageNumber			constant varchar2 (2)	:= '&P';		-- Номер страницы		&[Страница]
xl_hd_NumberOfPages			constant varchar2 (2)	:= '&N';		-- Количество страниц	&[Страниц]
xl_hd_CurDate				constant varchar2 (2)	:= '&D';		-- Текущая дата			&[Дата]
xl_hd_CurTime				constant varchar2 (2)	:= '&T';		-- Текущая время		&[Время]
xl_hd_Full					constant varchar2 (4)	:= '&Z&F';		-- Путь к файлу			&[Путь]&[Файл]
xl_hd_FileName				constant varchar2 (2)	:= '&F';		-- Имя файла			&[Файл]
xl_hd_SheetName				constant varchar2 (2)	:= '&A';		-- Имя листа			&[Лист]

type tp_alignment is record
    ( vertical varchar2(11),
      horizontal varchar2(16),
      wrapText boolean
    );        
    
-- Очистка локальной переменной, хранящей данные книги
  procedure clear_workbook;
-- Добавление в книгу нового листа
  procedure new_sheet
    ( p_sheetname varchar2 := null -- Наименование нового листа
    , p_tabcolor varchar2 := null  -- Цвет нового листа
    );
-- Преобразование маски формата для даты из Oracle в Excel
  function OraFmt2Excel( p_format varchar2 := null )
  return varchar2;
-- Получение идентификатора формата. 
  function get_numFmt( p_format varchar2 := null )
  return pls_integer;
-- Получение номера шрифта
  function get_font
    ( p_name varchar2              -- Наименование шрифта (например, 'Calibri')
    , p_family pls_integer := 2    -- строка значения для тэга <family val="..."/>
    , p_fontsize number := 11      -- размер шрифта
    , p_theme pls_integer := 1     -- тема - шаблон стилей
    , p_underline boolean := false -- признак у шрифта свойства подчёркивания
    , p_italic boolean := false    -- признак у шрифта свойства курсива
    , p_bold boolean := false      -- признак жирности шрифта
   , p_rgb varchar2 := null        -- 16-тиричное символьное значение Прозрачность Красный Зелёный Синий" (alpha red green blue) (например, 'FFFF0000')
    )
  return pls_integer;
-- Идентификатора для тэга <fills
  function get_fill
    ( p_patternType varchar2   -- Тип
    , p_fgRGB varchar2 := null -- 16-тиричное символьное значение Прозрачность Красный Зелёный Синий" (alpha red green blue) (например, 'FFFF0000')
    )
  return pls_integer;
-- Идентификатор для окантовок ячейки
  function get_border
    ( p_top varchar2      -- окантовка верхней линии
    , p_bottom varchar2   -- окантовка нижней линии
    , p_left varchar2     -- окантовка левой линии
    , p_right varchar2    -- окантовка правой линии
    )
  return pls_integer;
-- Выравнивание
  function get_alignment
    ( p_vertical varchar2 := null   -- По вертикали
    , p_horizontal varchar2 := null -- По горизонтали
    , p_wrapText boolean := null    -- Необходимость переноса данных в ячейке по словам
    )
  return tp_alignment;
-- Запись числового значения в ячейку
  procedure cell
    ( p_col pls_integer                   -- Колонка адреса ячейки
    , p_row pls_integer                   -- Строка адреса ячейки
    , p_value number                      -- Числовое значение для записи в ячейку
    , p_numFmtId pls_integer := null      -- Идентификатор формата
    , p_fontId pls_integer := null        -- Идентификатор фонта
    , p_fillId pls_integer := null        -- Идентификатор для тэга <fills />
    , p_borderId pls_integer := null      -- Идентификатор окантовок
    , p_alignment tp_alignment := null    -- Выравнивание
    , p_sheet pls_integer := null         -- Номер листа
    );
-- Запись строкового значения в ячейку
  procedure cell                          --
    ( p_col pls_integer                   -- Колонка адреса ячейки
    , p_row pls_integer                   -- Строка адреса ячейки
    , p_value varchar2                    -- Строковое значение для записи в ячейку
    , p_numFmtId pls_integer := null      -- Идентификатор формата
    , p_fontId pls_integer := null        -- Идентификатор фонта
    , p_fillId pls_integer := null        -- Идентификатор для тэга <fills />
    , p_borderId pls_integer := null      -- Идентификатор окантовок
    , p_alignment tp_alignment := null    -- Выравнивание
    , p_sheet pls_integer := null         -- Номер листа
    );
-- Запись значения типа дата в ячейку
  procedure cell
    ( p_col pls_integer                   -- Колонка адреса ячейки
    , p_row pls_integer                   -- Строка адреса ячейки
    , p_value date                        -- Значение типа дата для записи в ячейку
    , p_numFmtId pls_integer := null      -- Идентификатор формата
    , p_fontId pls_integer := null        -- Идентификатор фонта
    , p_fillId pls_integer := null        -- Идентификатор для тэга <fills />
    , p_borderId pls_integer := null      -- Идентификатор окантовок
    , p_alignment tp_alignment := null    -- Выравнивание
    , p_sheet pls_integer := null         -- Номер листа
    );
-- Работа с тэгом <Relationship Type=".../hyperlink" Target="P_URL" TargetMode="External"/>
  procedure hyperlink
    ( p_col pls_integer            -- Колонка адреса ячейки
    , p_row pls_integer            -- Строка адреса ячейки
    , p_url varchar2               -- Значение для атрибута Target="..."
    , p_value varchar2 := null     -- Строка в файле xl/sharedStrings.xml
    , p_sheet pls_integer := null  -- Номер листа
    );
-- Добавление комментария к ячейке
  procedure comment
    ( p_col pls_integer            -- Колонка адреса ячейки
    , p_row pls_integer            -- Строка адреса ячейки
    , p_text varchar2              -- Текст комментария
    , p_author varchar2 := null    -- Автор комментария
    , p_width pls_integer := 150   -- Ширина комментария в pixels
    , p_height pls_integer := 100  -- Высота комментария в pixels
    , p_sheet pls_integer := null  -- Номер листа
    );
-- Объединение ячеек
  procedure mergecells
    ( p_tl_col pls_integer          -- Колонка адреса верхней левой (top left) ячейки
    , p_tl_row pls_integer          -- Строка адреса верхней левой (top left) ячейки
    , p_br_col pls_integer          -- Колонка адреса нижней правой (bottom right) ячейки
    , p_br_row pls_integer          -- Строка адреса нижней правой (bottom right) ячейки
    , p_sheet pls_integer := null   -- Номер листа
    );
-- Сноски для диапозона (системная информация)
  procedure list_validation
    ( p_sqref_col pls_integer          -- Колонка адреса для значения в тэге <dataValidation />
    , p_sqref_row pls_integer          -- Строка адреса для значения в тэге <dataValidation />
    , p_tl_col pls_integer             -- Колонка адреса верхней левой (top left) ячейки
    , p_tl_row pls_integer             -- Строка адреса верхней левой (top left) ячейки
    , p_br_col pls_integer             -- Колонка адреса нижней правой (bottom right) ячейки
    , p_br_row pls_integer             -- Строка адреса нижней правой (bottom right) ячейки
    , p_style varchar2 := xl_inf_stop  -- Категория ошибки
    , p_title varchar2 := null         -- Текст для заголовка
    , p_prompt varchar := null         -- Текст для подсказки
    , p_show_error boolean := false    -- Необходимость отображения ошибок
    , p_error_title varchar2 := null   -- Текст для заголовка ошибки
    , p_error_txt varchar2 := null     -- Текст ошибки
    , p_sheet pls_integer := null      -- Номер листа
    );
-- Сноски (ValidAlert...) для диапозона (системная информация)
  procedure list_validation
    ( p_sqref_col pls_integer             -- Колонка адреса для значения в тэге <dataValidation />
    , p_sqref_row pls_integer             -- Строка адреса для значения в тэге <dataValidation />
    , p_defined_name varchar2             -- значение для тэга <formula1>
    , p_style varchar2 := xl_inf_stop     -- Категория ошибки
    , p_title varchar2 := null            -- Текст для заголовка
    , p_prompt varchar := null            -- Текст для подсказки
    , p_show_error boolean := false       -- Необходимость отображения ошибок
    , p_error_title varchar2 := null      -- Текст для заголовка ошибки
    , p_error_txt varchar2 := null        -- Текст ошибки
    , p_sheet pls_integer := null         -- Номер листа
    );
-- Присвоение имени диапазону ячеек
  procedure defined_name
    ( p_tl_col pls_integer                -- Колонка адреса верхней левой (top left) ячейки
    , p_tl_row pls_integer                -- Строка адреса верхней левой (top left) ячейки
    , p_br_col pls_integer                -- Колонка адреса нижней правой (bottom right) ячейки
    , p_br_row pls_integer                -- Строка адреса нижней правой (bottom right) ячейки
    , p_name varchar2                     -- Имя для диапазона ячеек
    , p_sheet pls_integer := null         -- Порядковый номер листа
    , p_localsheet pls_integer := null    -- Идентификатор листа
    );
-- Установка ширины колонки
  procedure set_column_width
    ( p_col pls_integer            -- номер колонки
    , p_width number               -- Значение ширины колонки
    , p_sheet pls_integer := null  -- Номер листа
    );
-- Установка свойств колонки
  procedure set_column
    ( p_col pls_integer                   -- Номер колонки
    , p_numFmtId pls_integer := null      -- Идентификатор формата
    , p_fontId pls_integer := null        -- Идентификатор фонта
    , p_fillId pls_integer := null        -- Идентификатор для тэга <fills />
    , p_borderId pls_integer := null      -- Идентификатор окантовок
    , p_alignment tp_alignment := null    -- Выравнивание
    , p_sheet pls_integer := null         -- Номер листа
    );
-- Установка свойств строки
  procedure set_row
    ( p_row pls_integer                   -- Номер строки
    , p_numFmtId pls_integer := null      -- Идентификатор формата
    , p_fontId pls_integer := null        -- Идентификатор фонта
    , p_fillId pls_integer := null        -- Идентификатор для тэга <fills />
    , p_borderId pls_integer := null      -- Идентификатор окантовок
    , p_alignment tp_alignment := null    -- Выравнивание
    , p_sheet pls_integer := null         -- Номер листа
    , p_height number := null             -- Высота
    );
-- Закрепить строки
  procedure freeze_rows
    ( p_nr_rows pls_integer := 1         -- Количество строк
    , p_sheet pls_integer := null        -- Номер листа
    );
-- Закрепить колонки
  procedure freeze_cols
    ( p_nr_cols pls_integer := 1         -- Количество колонок
    , p_sheet pls_integer := null        -- Номер листа
    );
-- Закрепить область
  procedure freeze_pane
    ( p_col pls_integer                  -- Количество колонок
    , p_row pls_integer                  -- Количество строк
    , p_sheet pls_integer := null        -- Номер листа
    );
-- Установить на строке автофильтр
  procedure set_autofilter
    ( p_column_start pls_integer := null  -- Начальная колонка адреса автофильтра
    , p_column_end pls_integer := null    -- Конечная колонка адреса автофильтра
    , p_row_start pls_integer := null     -- Начальная строка адреса автофильтра
    , p_row_end pls_integer := null       -- Конечная строка адреса автофильтра
    , p_sheet pls_integer := null         -- Номер листа
    );
-- Установить цвет листа
  procedure set_tabcolor
    ( p_tabcolor varchar2             -- 16-тиричное символьное значение Прозрачность Красный Зелёный Синий" (alpha red green blue) (например, 'FFFF0000')
    , p_sheet pls_integer := null     -- Номер листа
    );
-- Формирование по загруженным данным xlsx файла
  function finish
  return blob;      -- Возвращается xlsx файл
-- Сохранение сформированного файла в указанную директорию
  procedure save
    ( p_directory varchar2   -- Наименования директории
    , p_filename varchar2    -- Наименование файла
    );
-- Выгрузка Grid в Excel на сервер (упрощенная версия)
  procedure query2sheet
    ( p_sql varchar2                     -- sql запрос (Например 'select 1 from dual')
    , p_column_headers boolean := true   -- Выгружать ли заголовки?
    , p_directory varchar2 := null       -- Наименования директории
    , p_filename varchar2 := null        -- Наименование файла
    , p_sheet pls_integer := null        -- Номер листа
    , p_UseXf boolean := false           -- Использовать ли формат XF
    );
-- Использовать XF спецификацию 
  procedure setUseXf( p_val boolean := true ); 

--ВБРР  возвращает Excel в blob-переменной по sql-запросу
  function query2blob
    ( p_sql varchar2                     -- sql запрос (Например 'select 1 from dual')
    , p_column_headers boolean := true   -- Выгружать ли заголовки?
    , p_sheet pls_integer := null        -- Номер листа
    , p_UseXf boolean := false           -- Использовать ли формат XF
    ) return blob;  
  
end;

/
