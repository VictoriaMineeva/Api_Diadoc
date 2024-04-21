using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Diadoc.Api;
using Diadoc.Api.Proto.Invoicing;
using OfficeOpenXml;
            // Задаем логин, пароль и другие атрибуты для входа
            string login = "vikenti2010@mail.ru";
            string apiKey = "RATest-5347d365-7f18-4126-b5b6-e76ab0073563";

            Console.Write("Введите пароль: ");
            string password = Console.ReadLine();

            // Инициализация клиента Diadoc
            var diadocClient = new DiadocApi(apiKey);

            // Авторизуемся в системе Diadoc
            var authResult = await diadocClient.AuthenticateAsync(login, password);

            if (!authResult.IsAuthenticated)
            {
                Console.WriteLine("Не удалось авторизоваться. Проверьте логин, пароль и ключ API.");
                return;
            }

            // Получение списка организаций, привязанных к аккаунту
            var orgs = await diadocClient.GetOrganizationsAsync();

            // Создаем файл Excel и добавляем в него данные об организациях и сотрудниках
            string filePath = @"Excel/Excel.xlsx"; // Укажите свой путь к файлу
            CreateExcelFile(filePath, orgs, diadocClient);

            Console.WriteLine("Данные успешно сохранены в файл Excel.");
            Console.ReadLine();
        }

        // Метод для создания и заполнения файла Excel данными об организациях и сотрудниках
        static void CreateExcelFile(string filePath, List<OrganizationInfo> orgs, DiadocApi diadocClient)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("OrganizationsAndUsers");

                // Добавляем заголовки столбцов для организаций
                var orgHeaders = new string[] { "OrgIdGuid", "OrgId", "Inn", "Kpp", "FullName", "ShortName",
                                                 "JoinedDiadocTreaty", "FnsParticipantId", "Departments", "IsPilot",
                                                 "IsActive", "IsTest", "IsBranch", "IsRoaming", "IsEmployee",
                                                 "InvitationCount", "SearchCount", "Sociability", "IsForeign", "HasCertificateToSign" };

                // Добавляем заголовки столбцов для сотрудников
                var userHeaders = new string[] { "OrgId", "UserName", "UserId", "AuthorizationPermission", "CanAddResolutions",
                                                  "CanCreateDocuments", "CanDeleteRestoreDocuments", "CanManageCounteragents",
                                                  "CanRequestResolutions", "CanSendDocuments", "CanSignDocuments",
                                                  "DocumentAccessLevel", "IsAdministrator", "JobTitle",
                                                  "SelectedDepartmentIds", "UserDepartmentId", "Position", "CurrentUserId" };

                // Записываем заголовки столбцов
                for (int i = 0; i < orgHeaders.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = orgHeaders[i];
                }

                for (int i = 0; i < userHeaders.Length; i++)
                {
                    worksheet.Cells[1, i + orgHeaders.Length + 2].Value = userHeaders[i];
                }

                int row = 2;

                // Записываем данные об организациях и сотрудниках

                foreach (var org in orgs)
                {
                    // Записываем данные об организации
                    worksheet.Cells[row, 1].Value = org.Id;
                    worksheet.Cells[row, 2].Value = org.OrgId;
                    worksheet.Cells[row, 3].Value = org.Inn;
                    worksheet.Cells[row, 4].Value = org.Kpp;
                    worksheet.Cells[row, 5].Value = org.FullName;
                    worksheet.Cells[row, 6].Value = org.ShortName;
                    worksheet.Cells[row, 7].Value = org.JoinedDiadocTreaty;
                    worksheet.Cells[row, 8].Value = org.FnsParticipantId;
                    worksheet.Cells[row, 9].Value = org.Departments;
                    worksheet.Cells[row, 10].Value = org.IsPilot;
                    worksheet.Cells[row, 11].Value = org.IsActive;
                    worksheet.Cells[row, 12].Value = org.IsTest;
                    worksheet.Cells[row, 13].Value = org.IsBranch;
                    worksheet.Cells[row, 14].Value = org.IsRoaming;
                    worksheet.Cells[row, 15].Value = org.IsEmployee;
                    worksheet.Cells[row, 16].Value = org.InvitationCount;
                    worksheet.Cells[row, 17].Value = org.SearchCount;
                    worksheet.Cells[row, 18].Value = org.Sociability;
                    worksheet.Cells[row, 19].Value = org.IsForeign;
                    worksheet.Cells[row, 20].Value = org.HasCertificateToSign;

                    // Получаем список пользователей организации
                    var users = diadocClient.GetOrganizationUsers(org.Id).Result;

                    // Записываем данные о сотрудниках организации
                    foreach (var user in users)
                    {
                        worksheet.Cells[row, orgHeaders.Length + 2].Value = org.ShortName; // Добавляем название организации в каждую строку с сотрудниками
                        worksheet.Cells[row, orgHeaders.Length + 3].Value = user.Name;
                        worksheet.Cells[row, orgHeaders.Length + 4].Value = user.Id;
                        worksheet.Cells[row, orgHeaders.Length + 5].Value = user.Permissions.AuthorizationPermission;
                        worksheet.Cells[row, orgHeaders.Length + 6].Value = user.Permissions.CanAddResolutions;
                        worksheet.Cells[row, orgHeaders.Length + 7].Value = user.Permissions.CanCreateDocuments;
                        worksheet.Cells[row, orgHeaders.Length + 8].Value = user.Permissions.CanDeleteRestoreDocuments;
                        worksheet.Cells[row, orgHeaders.Length + 9].Value = user.Permissions.CanManageCounteragents;
                        worksheet.Cells[row, orgHeaders.Length + 10].Value = user.Permissions.CanRequestResolutions;
                        worksheet.Cells[row, orgHeaders.Length + 11].Value = user.Permissions.CanSendDocuments;
                        worksheet.Cells[row, orgHeaders.Length + 12].Value = user.Permissions.CanSignDocuments;
                        worksheet.Cells[row, orgHeaders.Length + 13].Value = user.Permissions.DocumentAccessLevel;
                        worksheet.Cells[row, orgHeaders.Length + 14].Value
.Value = user.Permissions.IsAdministrator;
                        worksheet.Cells[row, orgHeaders.Length + 15].Value = user.JobTitle;
                        worksheet.Cells[row, orgHeaders.Length + 16].Value = string.Join(",", user.SelectedDepartmentIds);
                        worksheet.Cells[row, orgHeaders.Length + 17].Value = user.UserDepartmentId;
                        worksheet.Cells[row, orgHeaders.Length + 18].Value = user.Position;
                        worksheet.Cells[row, orgHeaders.Length + 19].Value = user.CurrentUserId;

                        row++;
                    }
                }

                // Сохраняем файл Excel
                FileInfo excelFile = new FileInfo(filePath);
                package.SaveAs(excelFile);
            }
        }
