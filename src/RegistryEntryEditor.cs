﻿using System;
using System.Windows.Forms;

namespace RD_AAOW
	{
	/// <summary>
	/// Форма обеспечивает создание или редактирование записи в базе
	/// </summary>
	public partial class RegistryEntryEditor: Form
		{
		/// <summary>
		/// Возвращает изменённую запись
		/// </summary>
		public RegistryEntry EditedEntry
			{
			get
				{
				return editedEntry;
				}
			}
		private RegistryEntry editedEntry;

		/// <summary>
		/// Флаг указывает, что изменение записи было подтверждено
		/// </summary>
		public bool Confirmed
			{
			get
				{
				return confirmed;
				}
			}
		private bool confirmed = false;

		/// <summary>
		/// Конструктор. Запускает создание или редактирование записи
		/// </summary>
		/// <param name="Entry">Запись реестра</param>
		/// <param name="New">Флаг режима создания новой записи</param>
		public RegistryEntryEditor (RegistryEntry Entry, bool New)
			{
			// Инициализация
			InitializeComponent ();
			this.AcceptButton = Apply;
			this.CancelButton = Abort;

			// Сохранение параметров
			editedEntry = Entry;

			// Настройка контролов
			KeyType.Items.Add (RegistryValueTypes.REG_SZ.ToString () +
				RDLocale.GetText ("RegistryValue_String"));
			KeyType.Items.Add (RegistryValueTypes.REG_DWORD.ToString () +
				RDLocale.GetText ("RegistryValue_Int32"));
			KeyType.Items.Add (RegistryValueTypes.REG_QWORD.ToString () +
				RDLocale.GetText ("RegistryValue_Int64"));

			if ((int)editedEntry.ValueType < KeyType.Items.Count)
				KeyType.Text = KeyType.Items[(int)editedEntry.ValueType].ToString ();
			else
				KeyType.Text = KeyType.Items[0].ToString ();

			// Загрузка записи
			KeyPath.Text = editedEntry.ValuePath;
			KeyName.Text = editedEntry.ValueName;
			KeyObject.Text = editedEntry.ValueObject;
			PathMustBeDeleted.Checked = editedEntry.PathMustBeDeleted;
			NameMustBeDeleted.Checked = editedEntry.NameMustBeDeleted;

			// Локализация
			this.Text = RDLocale.GetText ("REE_Title");

			RDLocale.SetControlsText (this);
			Apply.Text = RDLocale.GetDefaultText (New ? RDLDefaultTexts.Button_Add :
				RDLDefaultTexts.Button_Apply);
			Abort.Text = RDLocale.GetDefaultText (RDLDefaultTexts.Button_Cancel);

			// Запуск
			this.ShowDialog ();
			}

		// Изменение состояния флага удаления раздела
		private void PathMustBeDeleted_CheckedChanged (object sender, EventArgs e)
			{
			KeyName.Enabled = KeyObject.Enabled = KeyType.Enabled = NameMustBeDeleted.Enabled =
				!PathMustBeDeleted.Checked;
			}

		// Отмена изменений
		private void Abort_Click (object sender, EventArgs e)
			{
			this.Close ();
			}

		// Применение изменений
		private void Apply_Click (object sender, EventArgs e)
			{
			// Проверка записи на корректность
			RegistryEntry re = new RegistryEntry (KeyPath.Text, KeyName.Text, KeyObject.Text,
				(RegistryValueTypes)KeyType.SelectedIndex, PathMustBeDeleted.Checked, NameMustBeDeleted.Checked);

			if (!re.IsValid)
				{
				RDInterface.LocalizedMessageBox (RDMessageTypes.Warning_Center, "EntryIsIncorrect");
				return;
				}

			// Успешно
			editedEntry = (RegistryEntry)re.Clone ();
			confirmed = true;
			this.Close ();
			}
		}
	}
