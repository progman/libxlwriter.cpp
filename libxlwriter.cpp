//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
// 0.0.1
// Alexey Potehin <gnuplanet@gmail.com>, http://www.gnuplanet.ru/doc/cv
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
#include <stdio.h>
#include <string.h>
#include <errno.h>

#include "libxlwriter.hpp"
#include "submodule/libcore.cpp/libcore.hpp"
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
/**
 * \file       libxlwriter.cpp
 * \Author     gnuplanet@gmail.com
 * \brief      Конструктор
 * \param [in] name имя вкладки в Excel документе
 * \param [in] row_count количество строк
 * \param [in] col_count количество колонок
 */
libxlwriter_sheet_t::libxlwriter_sheet_t(const std::string &name, uint16_t row_count, uint16_t col_count)
{
	this->name      = name;
	this->row_count = row_count;
	this->col_count = col_count;
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
/**
 * \file   libxlwriter.cpp
 * \Author gnuplanet@gmail.com
 * \brief  Деструктор
 */
libxlwriter_sheet_t::~libxlwriter_sheet_t()
{
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
/**
 * \file        libxlwriter.cpp
 * \Author      gnuplanet@gmail.com
 * \brief       Возвращает текст ошибки
 * \param [out] error_str текст ошибки
 */
void libxlwriter_sheet_t::get_error_str(std::string &error_str)
{
	error_str = this->error_str;
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
/**
 * \file       libxlwriter.cpp
 * \Author     gnuplanet@gmail.com
 * \brief      Задает тип и имя колонки
 * \param [in] col номер колонки
 * \param [in] sheet_type тип колонки
 * \param [in] имя колонки
 * \return     флаг успешности операции
 */
bool libxlwriter_sheet_t::set_col(uint16_t col, const libxlwriter_sheet_t::sheet_type_t sheet_type, const std::string headname)
{
	if (col >= this->col_count)
	{
		this->error_str = "invalid col number";
		return false;
	}


	libxlwriter_sheet_t::col_info_t col_info;
	col_info.sheet_type = sheet_type;
	col_info.headname   = headname;


	this->map_col[col] = col_info;


	return true;
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
/**
 * \file       libxlwriter.cpp
 * \Author     gnuplanet@gmail.com
 * \brief      Задает значение ячейки
 * \param [in] row номер строки
 * \param [in] col номер колонки
 * \param [in] значение ячейки
 * \return     флаг успешности операции
 */
bool libxlwriter_sheet_t::set(uint16_t row, uint16_t col, const std::string &data)
{
	if (row >= this->row_count)
	{
		this->error_str = "invalid row number";
		return false;
	}

	if (col >= this->col_count)
	{
		this->error_str = "invalid col number";
		return false;
	}

	uint32_t index = row;
	index <<= 16;
	index |= col;

	this->map_cell[index] = data;


	return true;
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
/**
 * \file   libxlwriter.cpp
 * \Author gnuplanet@gmail.com
 * \brief  Конструктор
 */
libxlwriter_t::libxlwriter_t()
{
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
/**
 * \file   libxlwriter.cpp
 * \Author gnuplanet@gmail.com
 * \brief  Деструктор
 */
libxlwriter_t::~libxlwriter_t()
{
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
/**
 * \file        libxlwriter.cpp
 * \Author      gnuplanet@gmail.com
 * \brief       Возвращает текст ошибки
 * \param [out] error_str текст ошибки
 */
void libxlwriter_t::get_error_str(std::string &error_str)
{
	error_str = this->error_str;
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
/**
 * \file       libxlwriter.cpp
 * \Author     gnuplanet@gmail.com
 * \brief      Добавляет вкладку в Excel документ
 * \param [in] sheet вкладка Excel документа
 */
void libxlwriter_t::add_sheet(const libxlwriter_sheet_t &sheet)
{
	this->sheet_list.push_back(sheet);
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
/**
 * \file        libxlwriter.cpp
 * \Author      gnuplanet@gmail.com
 * \brief       Получает содержимое Excel документа
 * \param [out] document содержимое Excel документа
 * \param [in]  flag_clear флаг определяющий стирать ли содержимое переменной прежде чем поместить в нее содержимое Excel документа
 * \return      флаг успешности операции
 */
bool libxlwriter_t::get(std::string &document, bool flag_clear)
{
	if (flag_clear != false)
	{
		document.clear();
	}


	document.append("<?xml version=\"1.0\"?>\n");
	document.append("<?mso-application progid=\"Excel.Sheet\"?>\n");
	document.append("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n\n");

	document.append("\t<Styles>\n\n");

	document.append("\t\t<Style ss:ID=\"s62\">\n");
	document.append("\t\t\t<NumberFormat ss:Format=\"Short Date\"/>\n");
	document.append("\t\t</Style>\n\n");

	document.append("\t\t<Style ss:ID=\"s63\">\n");
	document.append("\t\t\t<NumberFormat ss:Format=\"_-* #,##0.00&quot;р.&quot;_-;\\-* #,##0.00&quot;р.&quot;_-;_-* &quot;-&quot;??&quot;р.&quot;_-;_-@_-\"/>\n");
	document.append("\t\t</Style>\n\n");

	document.append("\t</Styles>\n\n");


	for (std::list<libxlwriter_sheet_t>::iterator i = this->sheet_list.begin(); i != this->sheet_list.end(); ++i)
	{
		if ((*i).row_count == 0)
		{
			this->error_str = "emoty row count";
			return false;
		}

		if ((*i).col_count == 0)
		{
			this->error_str = "emoty col count";
			return false;
		}


		bool flag_header = false;
		for (uint16_t col = 0; col < (*i).col_count; col++)
		{
			if ((*i).map_col.find(col) == (*i).map_col.end())
			{
				this->error_str = "col not found";
				return false;
			}
			libxlwriter_sheet_t::col_info_t col_info = (*i).map_col[col];
			if (col_info.headname.empty() == false)
			{
				flag_header = true;
				break;
			}
		}


		std::string col_str;
		if (libcore::uint2str(col_str, (*i).col_count) == false)
		{
			this->error_str = "invalid col number";
			return false;
		}

		std::string row_str;
		if (flag_header == false)
		{
			if (libcore::uint2str(row_str, (*i).row_count) == false)
			{
				this->error_str = "invalid row number";
				return false;
			}
		}
		else
		{
			if (libcore::uint2str(row_str, (*i).row_count + 1) == false)
			{
				this->error_str = "invalid row number";
				return false;
			}
		}


		document.append("\t<Worksheet ss:Name=\"" + (*i).name + "\">\n");
		document.append("\t\t<Table ss:ExpandedColumnCount=\"" + col_str + "\" ss:ExpandedRowCount=\"" + row_str + "\">\n\n");


		for (uint16_t col = 0; col < (*i).col_count; col++)
		{
			if ((*i).map_col.find(col) == (*i).map_col.end())
			{
				this->error_str = "row not found";
				return false;
			}
			libxlwriter_sheet_t::col_info_t col_info = (*i).map_col[col];

			switch (col_info.sheet_type)
			{
				case libxlwriter_sheet_t::DateTime:
				{
					document.append("\t\t\t<Column ss:AutoFitWidth=\"1\" ss:StyleID=\"s62\" />\n");
					break;
				}

				case libxlwriter_sheet_t::Number:
				{
					document.append("\t\t\t<Column ss:AutoFitWidth=\"1\" ss:StyleID=\"s63\" />\n");
					break;
				}

				case libxlwriter_sheet_t::String:
				{
					document.append("\t\t\t<Column ss:AutoFitWidth=\"1\" />\n");
					break;
				}

				default:
				{
					break;
				}
			}
		}
		document.append("\n");


		if (flag_header != false)
		{
			document.append("\t\t\t<Row ss:AutoFitHeight=\"0\">\n");
			for (uint16_t col = 0; col < (*i).col_count; col++)
			{
				if ((*i).map_col.find(col) == (*i).map_col.end())
				{
					this->error_str = "col not found";
					return false;
				}
				libxlwriter_sheet_t::col_info_t col_info = (*i).map_col[col];

				document.append("\t\t\t\t<Cell><Data ss:Type=\"String\">" + col_info.headname + "</Data></Cell>\n");
			}
			document.append("\t\t\t</Row>\n\n");
		}


		for (uint16_t row = 0; row < (*i).row_count; row++)
		{
			document.append("\t\t\t<Row ss:AutoFitHeight=\"0\">\n");
			for (uint16_t col = 0; col < (*i).col_count; col++)
			{
				if ((*i).map_col.find(col) == (*i).map_col.end())
				{
					this->error_str = "col not found";
					return false;
				}

				libxlwriter_sheet_t::col_info_t col_info = (*i).map_col[col];

				uint32_t index = row;
				index <<= 16;
				index |= col;

				if ((*i).map_cell.find(index) == (*i).map_cell.end())
				{
					this->error_str = "index not found";
					return false;
				}
				std::string data = (*i).map_cell[index];


				switch (col_info.sheet_type)
				{
					case libxlwriter_sheet_t::DateTime:
					{
						document.append("\t\t\t\t<Cell><Data ss:Type=\"DateTime\">" + data + "</Data></Cell>\n");
						break;
					}

					case libxlwriter_sheet_t::Number:
					{
						document.append("\t\t\t\t<Cell><Data ss:Type=\"Number\">" + data + "</Data></Cell>\n");
						break;
					}

					case libxlwriter_sheet_t::String:
					{
						document.append("\t\t\t\t<Cell><Data ss:Type=\"String\">" + data + "</Data></Cell>\n");
						break;
					}

					default:
					{
						break;
					}
				}
			}

			document.append("\t\t\t</Row>\n");
			document.append("\n");
		}

		document.append("\t\t</Table>\n");
		document.append("\t</Worksheet>\n");
		document.append("\n");
	}


	document.append("</Workbook>\n");


	return true;
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
/**
 * \file       libxlwriter.cpp
 * \Author     gnuplanet@gmail.com
 * \brief      Записывает в файл содержимое Excel документа
 * \param [in] filename имя файла
 * \param [in] flag_sync флаг определяющий сбрасывать ли файловый кеш после записи файла
 * \return     флаг успешности операции
 */
bool libxlwriter_t::write(const std::string &filename, bool flag_sync)
{
	int rc;


	std::string document;
	if (this->get(document) == false)
	{
		return false;
	}


	rc = libcore::file_set(filename.c_str(), document, flag_sync, true);
	if (rc == -1)
	{
		this->error_str = strerror(errno);
		return false;
	}


	return true;
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
