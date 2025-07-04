﻿namespace RD_AAOW
	{
	/// <summary>
	/// Результаты применения реестровой записи
	/// </summary>
	public enum RegistryEntryApplicationResults
		{
		/// <summary>
		/// Запись не применялась, или результат её применения полностью утерян
		/// </summary>
		NotApplied = 0,

		/// <summary>
		/// Запись применялась, но результат её применения частично утерян, или реестр частично соответствует записи
		/// </summary>
		PartiallyApplied = 1,

		/// <summary>
		/// Реестр полностью соответствует записи
		/// </summary>
		FullyApplied = 2,

		/// <summary>
		/// Не удаётся получить доступ к записи
		/// </summary>
		CannotGetAccess = -1
		}

	/// <summary>
	/// Типы данных реестра Windows
	/// </summary>
	public enum RegistryValueTypes
		{
		/// <summary>
		/// Строковый
		/// </summary>
		REG_SZ = 0,

		/// <summary>
		/// Целое число, 4 байта
		/// </summary>
		REG_DWORD = 1,

		/// <summary>
		/// Целое число, 8 байт
		/// </summary>
		REG_QWORD = 2,

		/// <summary>
		/// Двоичный
		/// </summary>
		REG_BINARY = 3,

		/// <summary>
		/// Строковый переменной длины
		/// </summary>
		REG_EXPAND_SZ = 4,

		/// <summary>
		/// Строковый с поддержкой абзацев
		/// </summary>
		REG_MULTI_SZ = 5,

		/// <summary>
		/// Сырые данные
		/// </summary>
		REG_NONE = 6
		}
	}
