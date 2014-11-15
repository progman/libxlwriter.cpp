//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
// 0.0.1
// Alexey Potehin <gnuplanet@gmail.com>, http://www.gnuplanet.ru/doc/cv
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
#ifndef XLWRITER__INCLUDED
#define XLWRITER__INCLUDED
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
#include <stdint.h>
#include <string>
#include <list>
#include <map>
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
/**
 * \file   libxlwriter.hpp
 * \Author gnuplanet@gmail.com
 * \brief  Содержит вкладку Excel документа
 */
class libxlwriter_sheet_t
{
	std::string name;
	uint16_t row_count;
	uint16_t col_count;
	std::map<uint32_t, std::string> map_cell;
	std::string error_str;

public:
	enum sheet_type_t { DateTime, Number, String };

/**
 * \file       libxlwriter.hpp
 * \Author     gnuplanet@gmail.com
 * \brief      Конструктор
 * \param [in] name имя вкладки в Excel документе
 * \param [in] row_count количество строк
 * \param [in] col_count количество колонок
 */
	libxlwriter_sheet_t(const std::string &name, uint16_t row_count, uint16_t col_count);

/**
 * \file   libxlwriter.hpp
 * \Author gnuplanet@gmail.com
 * \brief  Деструктор
 */
	~libxlwriter_sheet_t();

/**
 * \file        libxlwriter.hpp
 * \Author      gnuplanet@gmail.com
 * \brief       Возвращает текст ошибки
 * \param [out] error_str текст ошибки
 */
	void get_error_str(std::string &error_str);

/**
 * \file       libxlwriter.hpp
 * \Author     gnuplanet@gmail.com
 * \brief      Задает тип и имя колонки
 * \param [in] col номер колонки
 * \param [in] sheet_type тип колонки
 * \param [in] имя колонки
 * \return     флаг успешности операции
 */
	bool set_col(uint16_t col, const libxlwriter_sheet_t::sheet_type_t sheet_type, const std::string headname = "");

/**
 * \file       libxlwriter.hpp
 * \Author     gnuplanet@gmail.com
 * \brief      Задает значение ячейки
 * \param [in] row номер строки
 * \param [in] col номер колонки
 * \param [in] значение ячейки
 * \return     флаг успешности операции
 */
	bool set(uint16_t row, uint16_t col, const std::string &data);

private:
	struct col_info_t
	{
		libxlwriter_sheet_t::sheet_type_t sheet_type;
		std::string headname;

		col_info_t()
		{
		}

		col_info_t(const col_info_t &other)
		{
			this->sheet_type = other.sheet_type;
			this->headname   = other.headname;
		}

		col_info_t &operator=(const col_info_t &other)
		{
			this->sheet_type = other.sheet_type;
			this->headname   = other.headname;
			return *this;
		}
	};
	std::map<uint16_t, col_info_t> map_col;

	friend class libxlwriter_t;
};
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
/**
 * \file   libxlwriter.hpp
 * \Author gnuplanet@gmail.com
 * \brief  Содержит Excel документ
 */
class libxlwriter_t
{
	std::list<libxlwriter_sheet_t> sheet_list;
	std::string error_str;

public:
/**
 * \file   libxlwriter.hpp
 * \Author gnuplanet@gmail.com
 * \brief  Конструктор
 */
	libxlwriter_t();

/**
 * \file   libxlwriter.hpp
 * \Author gnuplanet@gmail.com
 * \brief  Деструктор
 */
	~libxlwriter_t();

/**
 * \file        libxlwriter.hpp
 * \Author      gnuplanet@gmail.com
 * \brief       Возвращает текст ошибки
 * \param [out] error_str текст ошибки
 */
	void get_error_str(std::string &error_str);

/**
 * \file       libxlwriter.hpp
 * \Author     gnuplanet@gmail.com
 * \brief      Добавляет вкладку в Excel документ
 * \param [in] sheet вкладка Excel документа
 */
	void add_sheet(const libxlwriter_sheet_t &sheet);

/**
 * \file        libxlwriter.hpp
 * \Author      gnuplanet@gmail.com
 * \brief       Получает содержимое Excel документа
 * \param [out] document содержимое Excel документа
 * \param [in]  flag_clear флаг определяющий стирать ли содержимое переменной прежде чем поместить в нее содержимое Excel документа
 * \return      флаг успешности операции
 */
	bool get(std::string &document, bool flag_clear = true);

/**
 * \file       libxlwriter.hpp
 * \Author     gnuplanet@gmail.com
 * \brief      Записывает в файл содержимое Excel документа
 * \param [in] filename имя файла
 * \param [in] flag_sync флаг определяющий сбрасывать ли файловый кеш после записи файла
 * \return     флаг успешности операции
 */
	bool write(const std::string &filename, bool flag_sync = false);
};
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
#endif
