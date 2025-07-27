using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace RD_AAOW
	{
	/// <summary>
	/// Главная форма программы
	/// </summary>
	public partial class FileAssociationsManagerForm: Form
		{
		// Переменные
		private List<RegistryEntriesBaseManager> rebm = [];
		private uint applied = 0, partiallyApplied = 0, notApplied = 0, noAccess = 0;

		/// <summary>
		/// Конструктор. Инициализирует главную форму приложения
		/// </summary>
		public FileAssociationsManagerForm ()
			{
			// Инициализация основных элементов
			InitializeComponent ();
			}

		private void MainForm_Load (object sender, EventArgs e)
			{
			// Настройка контролов
			this.Text = ProgramDescription.AssemblyTitle;
			RDGenerics.LoadWindowDimensions (this);

			LanguageCombo.Items.AddRange (RDLocale.LanguagesNames);
			try
				{
				LanguageCombo.SelectedIndex = (int)RDLocale.CurrentLanguage;
				}
			catch
				{
				LanguageCombo.SelectedIndex = 0;
				}

			MainTable.Columns.Add ("Entries", "Entries");
			MainTable.ContextMenuStrip = new ContextMenuStrip ();
			MainTable.ContextMenuStrip.ShowImageMargin = false;

			MainTable.ContextMenuStrip.Items.Add (EditRecord.Text, null, EditRecord_Click);
			MainTable.ContextMenuStrip.Items.Add (Apply.Text, null, Apply_Click);
			MainTable.ContextMenuStrip.Items.Add (DeleteRecord.Text, null, DeleteRecord_Click);

			// Миграция из FEM
			if (!RDGenerics.GetSettings ("MigrationDone", false))
				{
				RDGenerics.SetSettings ("MigrationDone", true);

				// Получение пути установки
				string femPath = RDGenerics.GetDPArrayRegistryValue ("FileExtensionsManager");
				if (string.IsNullOrWhiteSpace (femPath))
					goto control;

				// Получение списка
				int idx = femPath.IndexOf ('\t');
				femPath = femPath.Substring (0, idx) + "\\REBases";
				string[] files = RegistryEntriesBaseManager.GetFASets (femPath);

				// Копирование файлов
				for (int i = 0; i < files.Length; i++)
					{
					try
						{
						File.Copy (files[i], RDGenerics.AppStartupPath + RegistryEntriesBaseManager.BasesSubdirectory +
							"\\" + Path.GetFileName (files[i]));
						}
					catch { }
					}
				}

			// Инициализация баз реестровых записей
			if (Directory.Exists (RDGenerics.AppStartupPath + RegistryEntriesBaseManager.BasesSubdirectory))
				{
				string[] files = RegistryEntriesBaseManager.GetFASets ();

				for (int i = 0; i < files.Length; i++)
					{
					RegistryEntriesBaseManager re =
						new RegistryEntriesBaseManager (Path.GetFileNameWithoutExtension (files[i]), false);
					if (re.IsInited)
						rebm.Add (re);
					}
				}

			// Контроль
			control:
			if (rebm.Count == 0)
				{
				if (!AddBaseMethod ())
					{
					this.Close ();
					return;
					}
				}

			// Загрузка списка
			for (int i = 0; i < rebm.Count; i++)
				BasesCombo.Items.Add (rebm[i].BaseName);
			BasesCombo.SelectedIndex = 0;

			// Обновление таблицы
			UpdateTable ();
			}

		// Изменение размера окна
		private void MainForm_Resize (object sender, EventArgs e)
			{
			MainTable.Width = this.Width - 32;
			MainTable.Height = this.Height - 196;

			ButtonsPanel.Top = this.Height - 147;
			ButtonsPanel.Left = (this.Width - ButtonsPanel.Width) / 2 - 4;
			}

		// Обновление таблицы
		private void UpdateTable ()
			{
			// Запрос
			List<string> presentation = rebm[BasesCombo.SelectedIndex].EntriesBasePresentation;
			List<RegistryEntryApplicationResults> statuses = rebm[BasesCombo.SelectedIndex].EntriesStatusesPresentation;

			// Обновление
			applied = 0;
			partiallyApplied = 0;
			notApplied = 0;
			noAccess = 0;
			MainTable.Rows.Clear ();

			for (int i = 0; i < presentation.Count; i++)
				{
				MainTable.Rows.Add ();
				MainTable.Rows[i].Cells[0].Value = presentation[i];
				switch (statuses[i])
					{
					case RegistryEntryApplicationResults.CannotGetAccess:
						MainTable.Rows[i].DefaultCellStyle.BackColor = NoAccess.ForeColor;
						MainTable.Rows[i].DefaultCellStyle.SelectionBackColor = NoAccess.BackColor;
						noAccess++;
						break;

					case RegistryEntryApplicationResults.FullyApplied:
						MainTable.Rows[i].DefaultCellStyle.BackColor = Applied.ForeColor;
						MainTable.Rows[i].DefaultCellStyle.SelectionBackColor = Applied.BackColor;
						applied++;
						break;

					case RegistryEntryApplicationResults.PartiallyApplied:
						MainTable.Rows[i].DefaultCellStyle.BackColor = PartiallyApplied.ForeColor;
						MainTable.Rows[i].DefaultCellStyle.SelectionBackColor = PartiallyApplied.BackColor;
						partiallyApplied++;
						break;

					case RegistryEntryApplicationResults.NotApplied:
						MainTable.Rows[i].DefaultCellStyle.BackColor = NotApplied.ForeColor;
						MainTable.Rows[i].DefaultCellStyle.SelectionBackColor = NotApplied.BackColor;
						notApplied++;
						break;
					}
				}

			// Результаты
			UpdateResults ();
			}

		private void UpdateResults ()
			{
			Applied.Text = RDLocale.GetText ("AppliedText") + applied.ToString ();
			PartiallyApplied.Text = RDLocale.GetText ("PartiallyAppliedText") + partiallyApplied.ToString ();
			NotApplied.Text = RDLocale.GetText ("NotAppliedText") + notApplied.ToString ();
			NoAccess.Text = RDLocale.GetText ("NoAccessText") + noAccess.ToString ();
			}

		// Выход из программы
		private void Exit_Click (object sender, EventArgs e)
			{
			this.Close ();
			}

		private void FileAssociationsManagerForm_FormClosing (object sender, FormClosingEventArgs e)
			{
			RDMessageButtons res = RDInterface.LocalizedMessageBox (RDMessageFlags.Question | RDMessageFlags.CenterText,
				"SaveBasesMessage", RDLDefaultTexts.Button_Yes, RDLDefaultTexts.Button_No,
				RDLDefaultTexts.Button_Cancel);

			// Сохранение баз
			if (res == RDMessageButtons.ButtonOne)
				{
				for (int i = 0; i < rebm.Count; i++)
					rebm[i].SaveBase (true);
				}

			// Выход
			RDGenerics.SaveWindowDimensions (this);
			e.Cancel = (res == RDMessageButtons.ButtonThree);
			}

		// Удаление записи
		private void DeleteRecord_Click (object sender, EventArgs e)
			{
			// Контроль
			if (MainTable.SelectedRows.Count <= 0)
				return;

			if (RDInterface.LocalizedMessageBox (RDMessageFlags.Question | RDMessageFlags.CenterText, "RemoveEntry",
				RDLDefaultTexts.Button_YesNoFocus, RDLDefaultTexts.Button_No) != RDMessageButtons.ButtonOne)
				return;

			// Удаление
			List<int> idx = [];
			foreach (DataGridViewRow r in MainTable.SelectedRows)
				idx.Add (r.Index);
			idx.Sort ();

			for (int i = 0; i < idx.Count; i++)
				rebm[BasesCombo.SelectedIndex].DeleteEntry ((uint)(idx[i] - i));

			// Обновление таблицы
			UpdateTable ();
			if (rebm[BasesCombo.SelectedIndex].EntriesCount > 0)
				{
				if (rebm[BasesCombo.SelectedIndex].EntriesCount <= idx[idx.Count - 1])
					idx[idx.Count - 1] = (int)rebm[BasesCombo.SelectedIndex].EntriesCount - 1;

				MainTable.CurrentCell = MainTable.Rows[idx[idx.Count - 1]].Cells[0];
				}
			}

		// Применение выбранной записи
		private void Apply_Click (object sender, EventArgs e)
			{
			// Контроль
			if (MainTable.SelectedRows.Count <= 0)
				return;

			if (RDInterface.LocalizedMessageBox (RDMessageFlags.Question | RDMessageFlags.CenterText,
				"ApplyEntry", RDLDefaultTexts.Button_Yes, RDLDefaultTexts.Button_No) !=
				RDMessageButtons.ButtonOne)
				return;

			// Применение записей
			for (int i = 0; i < MainTable.SelectedRows.Count; i++)
				{
				string msg = "";
				switch (rebm[BasesCombo.SelectedIndex].GetRegistryEntry
					((uint)MainTable.SelectedRows[i].Index).ApplyEntry ())
					{
					case RegistryEntryApplicationResults.CannotGetAccess:
						msg = RDLocale.GetText ("EntryIsUnavailable");
						break;

					case RegistryEntryApplicationResults.PartiallyApplied:
					case RegistryEntryApplicationResults.NotApplied:
						msg = RDLocale.GetText ("EntryNotApplied");
						break;
					}

				if (!string.IsNullOrWhiteSpace (msg))
					RDInterface.MessageBox (RDMessageFlags.Warning | RDMessageFlags.CenterText, msg);
				}

			// Обновление таблицы
			int row = MainTable.SelectedRows[MainTable.SelectedRows.Count - 1].Index;
			UpdateTable ();
			MainTable.CurrentCell = MainTable.Rows[row].Cells[0];
			}

		// Применение всех записей
		private void ApplyAll_Click (object sender, EventArgs e)
			{
			// Контроль
			if (RDInterface.LocalizedMessageBox (RDMessageFlags.Warning | RDMessageFlags.CenterText,
				"ApplyAllEntries", RDLDefaultTexts.Button_YesNoFocus, RDLDefaultTexts.Button_No) !=
				RDMessageButtons.ButtonOne)
				return;

			if (RDInterface.LocalizedMessageBox (RDMessageFlags.Warning | RDMessageFlags.CenterText,
				"ApplyAllEntries2", RDLDefaultTexts.Button_Yes, RDLDefaultTexts.Button_No) !=
				RDMessageButtons.ButtonTwo)
				return;

			// Применение записей
			uint res = rebm[BasesCombo.SelectedIndex].ApplyAllEntries ();

			RDInterface.MessageBox (RDMessageFlags.Success | RDMessageFlags.CenterText,
				string.Format (RDLocale.GetText ("EntriesApplied"), res,
				rebm[BasesCombo.SelectedIndex].EntriesCount));

			// Обновление таблицы
			int row = 0;
			if (MainTable.SelectedRows.Count > 0)
				row = MainTable.SelectedRows[0].Index;

			UpdateTable ();
			if (MainTable.SelectedRows.Count > 0)
				MainTable.CurrentCell = MainTable.Rows[row].Cells[0];
			}

		// Загрузка из файла реестра
		private void LoadRegFile_Click (object sender, EventArgs e)
			{
			// Контроль
			if (OFDialog.ShowDialog () != DialogResult.OK)
				return;

			// Загрузка
			uint res = rebm[BasesCombo.SelectedIndex].LoadRegistryFile (OFDialog.FileName);

			RDInterface.MessageBox (RDMessageFlags.Success | RDMessageFlags.CenterText,
				RDLocale.GetText ("EntriesAdded") + res.ToString ());

			// Обновление таблицы
			int row = 0;
			if (MainTable.SelectedRows.Count > 0)
				row = MainTable.SelectedRows[0].Index;

			UpdateTable ();
			if (MainTable.SelectedRows.Count > 0)
				MainTable.CurrentCell = MainTable.Rows[row].Cells[0];
			}

		// Сохранение в файл реестра
		private void SaveRegFile_Click (object sender, EventArgs e)
			{
			// Контроль
			if (MainTable.SelectedRows.Count <= 0)
				return;

			// Извлечение выборки
			List<int> idx = [];
			foreach (DataGridViewRow r in MainTable.SelectedRows)
				idx.Add (r.Index);
			idx.Sort ();

			// Запрос сохранения
			if (SFDialog.ShowDialog () != DialogResult.OK)
				return;

			// Выполнение
			if (!rebm[BasesCombo.SelectedIndex].SaveRegistryFile (SFDialog.FileName, idx))
				RDInterface.MessageBox (RDMessageFlags.Warning | RDMessageFlags.CenterText,
					string.Format (RDLocale.GetDefaultText (RDLDefaultTexts.Message_SaveFailure_Fmt),
					SFDialog.FileName));
			}

		// Добавление записи
		private void AddRecord_Click (object sender, EventArgs e)
			{
			// Добавление
			int row = 0;
			RegistryEntryEditor ree;

			if (MainTable.SelectedRows.Count > 0)
				{
				row = MainTable.SelectedRows[0].Index;
				ree = new RegistryEntryEditor (rebm[BasesCombo.SelectedIndex].GetRegistryEntry ((uint)row), true);
				}
			else
				{
				ree = new RegistryEntryEditor (new RegistryEntry ("HKEY_CLASSES_ROOT\\", "", ""), true);
				}

			if (ree.Confirmed)
				{
				rebm[BasesCombo.SelectedIndex].AddEntry (ree.EditedEntry);

				// Обновление таблицы
				UpdateTable ();
				if (MainTable.SelectedRows.Count > 0)
					MainTable.CurrentCell = MainTable.Rows[row].Cells[0];
				}
			}

		// Редактирование записей
		private void EditRecord_Click (object sender, EventArgs e)
			{
			EditRecord_Click (null, null);
			}

		private void EditRecord_Click (object sender, DataGridViewCellEventArgs e)
			{
			// Контроль
			if (MainTable.SelectedRows.Count <= 0)
				return;

			// Редактирование
			int row = MainTable.SelectedRows[0].Index;
			RegistryEntryEditor ree = new RegistryEntryEditor
				(rebm[BasesCombo.SelectedIndex].GetRegistryEntry ((uint)row), false);
			if (!ree.Confirmed)
				return;

			rebm[BasesCombo.SelectedIndex].DeleteEntry ((uint)row);
			rebm[BasesCombo.SelectedIndex].AddEntry (ree.EditedEntry);

			// Обновление таблицы
			UpdateTable ();
			MainTable.CurrentCell = MainTable.Rows[row].Cells[0];

			// Запрос на применение
			Apply_Click (null, null);
			}

		private void MainTable_KeyDown (object sender, KeyEventArgs e)
			{
			switch (e.KeyCode)
				{
				case Keys.Return:
					EditRecord_Click (null, null);
					break;

				case Keys.Insert:
					AddRecord_Click (null, null);
					break;

				case Keys.Delete:
					DeleteRecord_Click (null, null);
					break;
				}
			}

		// Выбор текущей базы
		private void BasesCombo_SelectedIndexChanged (object sender, EventArgs e)
			{
			UpdateTable ();
			}

		// Добавление базы
		private void AddBase_Click (object sender, EventArgs e)
			{
			AddBaseMethod ();
			}

		private bool AddBaseMethod ()
			{
			// Запрос названия
			string name = RDInterface.MessageBox (RDLocale.GetText ("NewSetName"), true, 20);
			if (string.IsNullOrWhiteSpace (name))
				return false;

			// Попытка создания
			RegistryEntriesBaseManager re = new RegistryEntriesBaseManager (name, true);
			if (!re.IsInited)
				{
				RDInterface.LocalizedMessageBox (RDMessageFlags.Warning, "NewBaseNotAdded");
				return false;
				}

			// Успешно
			RDInterface.LocalizedMessageBox (RDMessageFlags.Success | RDMessageFlags.CenterText,
				"NewBaseAdded", 1000);
			rebm.Add (re);

			BasesCombo.Items.Add (re.BaseName);
			BasesCombo.SelectedIndex = BasesCombo.Items.Count - 1;

			return true;
			}

		// Запрос справки
		private void GetHelp_Click (object sender, EventArgs e)
			{
			RDInterface.ShowAbout (false);
			}

		// Вызов меню
		private void MainTable_CellContextClick (object sender, DataGridViewCellMouseEventArgs e)
			{
			if (e.Button == MouseButtons.Right)
				MainTable.ContextMenuStrip.Show ((Control)sender, new Point (e.X,
					e.Y + (e.RowIndex - MainTable.FirstDisplayedScrollingRowIndex) * MainTable.RowTemplate.Height));
			}

		// Просмотр иконок
		private void FindIcon_Click (object sender, EventArgs e)
			{
			IconsExtractor ie = new IconsExtractor ();

			if (ie.SelectedIconNumber >= 0)
				RDGenerics.SendToClipboard (ie.SelectedIconFile + "," + ie.SelectedIconNumber.ToString (), true);
			}

		// Регистрация расширения
		private void RegExtension_Click (object sender, EventArgs e)
			{
			ExtensionRegistrator er = new ExtensionRegistrator (rebm[BasesCombo.SelectedIndex]);
			if (er.Confirmed)
				UpdateTable ();
			}

		// Локализация формы
		private void LanguageCombo_SelectedIndexChanged (object sender, EventArgs e)
			{
			// Сохранение языка
			RDLocale.CurrentLanguage = (RDLanguages)LanguageCombo.SelectedIndex;

			// Локализация
			OFDialog.Title = SFDialog.Title = RDLocale.GetText ("FEMF_OFDialogTitle");
			OFDialog.Filter = SFDialog.Filter = RDLocale.GetText ("FEMF_OFDialogFilter");

			RDLocale.SetControlsText (this);
			RDLocale.SetControlsText (ButtonsPanel);
			AddRecord.Text = RDLocale.GetDefaultText (RDLDefaultTexts.Button_Add);
			EditRecord.Text = RDLocale.GetDefaultText (RDLDefaultTexts.Button_Edit);
			Exit.Text = RDLocale.GetDefaultText (RDLDefaultTexts.Button_Exit);

			UpdateResults ();
			}
		}
	}
