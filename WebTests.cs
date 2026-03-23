using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebUI
{
    public class WebTests
    {
        IWebDriver _webDriver = new EdgeDriver();

        [Fact]
        public void CorrectMainTitle()
        {
            _webDriver.Url = "https://test.webmx.ru/";
            const string Title = "Сервис заметок";
            Assert.Equal(Title, _webDriver.Title);
            _webDriver.Close();
        }

        [Fact]
        public void SwitchingModes()
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string currentClass = "tab active";

            IWebElement buttonEnter = _webDriver.FindElement(By.Id("loginTab"));
            IWebElement buttonReg = _webDriver.FindElement(By.Id("registerTab"));

            buttonReg.Click();

            Assert.Equal(currentClass, buttonReg.GetAttribute("class"));

            buttonEnter.Click();

            Assert.Equal(currentClass, buttonEnter.GetAttribute("class"));

            _webDriver.Close();
        }

        [Fact] 
        public void LogWithCorrectData()
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string classSection = "card auth-card";

            const string loginUser = "one1";
            const string passwordUser = "111111";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            IWebElement section = _webDriver.FindElement(By.Id("authSection"));

            Assert.NotEqual(classSection, section.GetAttribute("class"));

            const string welcome = $"Здравствуйте, {loginUser}!";

            IWebElement span = _webDriver.FindElement(By.Id("welcomeText"));

            Assert.Equal(welcome, span.GetAttribute("textContent"));

            _webDriver.Close();
        }

        [Fact]
        public void Message_WhiteSpaceLoginCorrectData()
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string textMessage = "Логин должен быть не короче 3 символов.";

            const string loginUser = "      ";
            const string passwordUser = "111111";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            IWebElement buttonReg = _webDriver.FindElement(By.Id("registerTab"));

            buttonReg.Click();

            Thread.Sleep(2000);

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            const string xPathMessage = "//*[@id=\"message\"]/span";

            IWebElement message = _webDriver.FindElement(By.XPath(xPathMessage));

            Assert.Equal(textMessage, message.GetAttribute("textContent"));

            _webDriver.Close();
        }

        [Theory]
        [InlineData("", "Заполните это поле.")] // В яндексе другие текста
        [InlineData("12", "Увеличьте текст до 3 симв. или более (в настоящее время используется 2 симв.).")]
        public void LoginValidation(string loginUser, string message)
        {
            _webDriver.Url = "https://test.webmx.ru/";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            string validation = inputLogin.GetAttribute("validationMessage");

            Assert.Equal(message, validation);

            _webDriver.Close();
        }

        [Theory]
        [InlineData("", "Заполните это поле.")] // В яндексе другие текста
        [InlineData("        ", "Заполните это поле.")]
        [InlineData("11111", "Увеличьте текст до 6 симв. или более (в настоящее время используется 5 симв.).")]
        public void PasswordValidation(string passwordUser, string message)
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string loginUser = "one1";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            string validation = inputPassword.GetAttribute("validationMessage");

            Assert.Equal(message, validation);

            _webDriver.Close();
        }

        [Theory]
        [InlineData("one1", "111112312")]
        [InlineData("0ne1", "111111")]
        public void Message_AboutIncorrectDataDuringLog(string loginUser, string passwordUser)
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string textMessage = "Неверный логин или пароль.";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            const string xPathMessage = "//*[@id=\"message\"]/span";

            IWebElement message = _webDriver.FindElement(By.XPath(xPathMessage));

            Assert.Equal(textMessage, message.GetAttribute("textContent"));

            _webDriver.Close();
        }

        [Fact]
        public void Message_RegisterAnExistingUser()
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string textMessage = "Пользователь с таким логином уже существует.";

            IWebElement buttonReg = _webDriver.FindElement(By.Id("registerTab"));

            buttonReg.Click();

            const string loginUser = "beta";
            const string passwordUser = "123456";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            const string xPathMessage = "//*[@id=\"message\"]/span";

            IWebElement message = _webDriver.FindElement(By.XPath(xPathMessage));

            Assert.Equal(textMessage, message.GetAttribute("textContent"));

            _webDriver.Close();
        }

        [Fact]
        public void Message_LogOutOfSystem()
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string classSection = "card auth-card";
            const string textMessage = "Вы вышли из системы.";

            const string loginUser = "one1";
            const string passwordUser = "111111";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            IWebElement logoutBtn = _webDriver.FindElement(By.Id("logoutBtn"));

            logoutBtn.Click();

            Thread.Sleep(2000);

            IWebElement section = _webDriver.FindElement(By.Id("authSection"));

            Assert.Equal(classSection, section.GetAttribute("class"));

            const string xPathMessage = "//*[@id=\"message\"]/span";

            IWebElement message = _webDriver.FindElement(By.XPath(xPathMessage));

            Assert.Equal(textMessage, message.GetAttribute("textContent"));

            _webDriver.Close();
        }

        [Theory]
        [InlineData("text")]
        [InlineData("beta")]
        [InlineData("Заметка1")]
        public void Message_AddEntry(string title)
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string textMessage = "Заметка создана.";

            const string loginUser = "one1";
            const string passwordUser = "111111";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            IWebElement newNoteBtn = _webDriver.FindElement(By.Id("newNoteBtn"));

            newNoteBtn.Click();

            IWebElement inputTitle = _webDriver.FindElement(By.Id("noteTitle"));
            inputTitle.SendKeys(title);

            IWebElement saveBtn = _webDriver.FindElement(By.Id("saveBtn"));

            saveBtn.Click();

            Thread.Sleep(2000);

            const string xPathMessage = "//*[@id=\"message\"]/span";

            IWebElement message = _webDriver.FindElement(By.XPath(xPathMessage));

            Assert.Equal(textMessage, message.GetAttribute("textContent"));

            const string xPathLi = "//*[@id=\"notesList\"]/li[1]/strong";
            IWebElement li = _webDriver.FindElement(By.XPath(xPathLi));

            Assert.Equal(title, li.GetAttribute("textContent"));

            _webDriver.Close();
        }

        [Fact]
        public void Message_DeleteUserEntry()
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string textMessage = "Заметка удалена.";

            const string loginUser = "one1";
            const string passwordUser = "111111";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            const string xPathLiStrong = "//*[@id=\"notesList\"]/li[1]/strong";
            IWebElement li_strong = _webDriver.FindElement(By.XPath(xPathLiStrong));

            string strong = li_strong.GetAttribute("textContent");

            IWebElement filter = _webDriver.FindElement(By.Id("noteScopeFilter"));

            filter.SendKeys("Мои");

            Thread.Sleep(2000);

            const string xPathLi = "//*[@id=\"notesList\"]/li";
            IWebElement li = _webDriver.FindElement(By.XPath(xPathLi));

            li.Click();

            IWebElement deleteBtn = _webDriver.FindElement(By.Id("deleteBtn"));

            deleteBtn.Click();

            var alert_win = _webDriver.SwitchTo().Alert();

            alert_win.Accept();

            Thread.Sleep(2000);

            const string xPathMessage = "//*[@id=\"message\"]/span";

            IWebElement message = _webDriver.FindElement(By.XPath(xPathMessage));

            Assert.Equal(textMessage, message.GetAttribute("textContent"));

            li_strong = _webDriver.FindElement(By.XPath(xPathLiStrong));

            Assert.NotEqual(strong, li_strong.GetAttribute("textContent"));

            _webDriver.Close();
        }

        [Fact]
        public void Message_CancelDeleteUserEntry()
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string loginUser = "one1";
            const string passwordUser = "111111";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            const string xPathLiStrong = "//*[@id=\"notesList\"]/li[1]/strong";
            IWebElement li_strong = _webDriver.FindElement(By.XPath(xPathLiStrong));

            string strong = li_strong.GetAttribute("textContent");

            IWebElement filter = _webDriver.FindElement(By.Id("noteScopeFilter"));

            filter.SendKeys("Мои");

            Thread.Sleep(2000);

            const string xPathLi = "//*[@id=\"notesList\"]/li";
            IWebElement li = _webDriver.FindElement(By.XPath(xPathLi));

            li.Click();

            IWebElement deleteBtn = _webDriver.FindElement(By.Id("deleteBtn"));

            deleteBtn.Click();

            var alert_win = _webDriver.SwitchTo().Alert();

            alert_win.Dismiss();

            Thread.Sleep(2000);

            li_strong = _webDriver.FindElement(By.XPath(xPathLiStrong));

            Assert.Equal(strong, li_strong.GetAttribute("textContent"));

            _webDriver.Close();
        }

        [Fact]
        public void Message_SaveChangesUserEntry()
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string textMessage = "Заметка обновлена.";
            const string titleBlock = "Новая заметка";

            const string loginUser = "one1";
            const string passwordUser = "111111";

            const string text = "two";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            IWebElement filter = _webDriver.FindElement(By.Id("noteScopeFilter"));

            filter.SendKeys("Мои");

            Thread.Sleep(2000);

            const string xPathLi = "//*[@id=\"notesList\"]/li";
            IWebElement li = _webDriver.FindElement(By.XPath(xPathLi));

            li.Click();

            IWebElement h2Title = _webDriver.FindElement(By.Id("editorTitle"));

            Assert.NotEqual(titleBlock, h2Title.GetAttribute("textContent"));

            IWebElement noteContent = _webDriver.FindElement(By.Id("noteContent"));
            noteContent.SendKeys(text);

            IWebElement saveBtn = _webDriver.FindElement(By.Id("saveBtn"));

            saveBtn.Click();

            Thread.Sleep(2000);

            const string xPathMessage = "//*[@id=\"message\"]/span";

            IWebElement message = _webDriver.FindElement(By.XPath(xPathMessage));

            Assert.Equal(textMessage, message.GetAttribute("textContent"));

            _webDriver.Close();
        }

        [Theory]
        [InlineData("gg", "gg")]
        [InlineData("w", "hj_newItem")]
        [InlineData("Что-нибудь", "hj_newItem")]
        [InlineData("de", "gg")]
        public void FindEntry(string entry, string result)
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string loginUser = "one1";
            const string passwordUser = "111111";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            IWebElement searchInput = _webDriver.FindElement(By.Id("searchInput"));

            searchInput.SendKeys(entry);

            Thread.Sleep(2000);

            const string xPathLiStrong = "//*[@id=\"notesList\"]/li[1]/strong";
            IWebElement li_strong = _webDriver.FindElement(By.XPath(xPathLiStrong));

            Assert.Equal(result, li_strong.GetAttribute("textContent"));

            _webDriver.Close();
        }

        [Fact]
        public void NullList()
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string titleEntry = "--";
            const string classLi = "empty";

            const string loginUser = "one1";
            const string passwordUser = "111111";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(5000);

            IWebElement searchInput = _webDriver.FindElement(By.Id("searchInput"));

            searchInput.SendKeys(titleEntry);

            Thread.Sleep(5000);

            const string xPathLi = "//*[@id=\"notesList\"]/li";
            IWebElement li = _webDriver.FindElement(By.XPath(xPathLi));

            Assert.Equal(classLi, li.GetAttribute("class"));

            _webDriver.Close();
        }

        [Fact]
        public void ValidationTitleEntry()
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string loginUser = "one1";
            const string passwordUser = "111111";

            const string message = "Заполните это поле.";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            IWebElement newNoteBtn = _webDriver.FindElement(By.Id("newNoteBtn"));

            newNoteBtn.Click();

            IWebElement inputTitle = _webDriver.FindElement(By.Id("noteTitle"));
            inputTitle.SendKeys("");

            IWebElement saveBtn = _webDriver.FindElement(By.Id("saveBtn"));

            saveBtn.Click();

            Thread.Sleep(2000);

            string validation = inputTitle.GetAttribute("validationMessage");

            Assert.Equal(message, validation);

            _webDriver.Close();
        }

        [Fact]
        public void Message_GiveAccessUser()
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string textMessage = "Доступ успешно выдан.";

            const string loginUser = "one1";
            const string passwordUser = "111111";
            const string user = "deamk2404";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            IWebElement filter = _webDriver.FindElement(By.Id("noteScopeFilter"));

            filter.SendKeys("Мои");

            Thread.Sleep(2000);

            const string xPathLi = "//*[@id=\"notesList\"]/li";
            IWebElement li = _webDriver.FindElement(By.XPath(xPathLi));

            li.Click();

            IWebElement shareUsername = _webDriver.FindElement(By.Id("shareUsername"));
            shareUsername.SendKeys(user);

            IWebElement shareBtn = _webDriver.FindElement(By.Id("shareBtn"));

            shareBtn.Click();

            Thread.Sleep(2000);

            const string xPathMessage = "//*[@id=\"message\"]/span";

            IWebElement message = _webDriver.FindElement(By.XPath(xPathMessage));

            Assert.Equal(textMessage, message.GetAttribute("textContent"));

            _webDriver.Close();
        }

        [Theory]
        [InlineData("", "Укажите логин пользователя для совместного доступа.")]
        [InlineData("aaaaaaaaaaaaaaaaa", "Пользователь не найден.")]
        public void Message_GiveAccessUser_IncorrectData(string user, string textMessage)
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string loginUser = "one1";
            const string passwordUser = "111111";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            IWebElement filter = _webDriver.FindElement(By.Id("noteScopeFilter"));

            filter.SendKeys("Мои");

            Thread.Sleep(2000);

            const string xPathLi = "//*[@id=\"notesList\"]/li";
            IWebElement li = _webDriver.FindElement(By.XPath(xPathLi));

            li.Click();

            IWebElement shareUsername = _webDriver.FindElement(By.Id("shareUsername"));
            shareUsername.SendKeys(user);

            IWebElement shareBtn = _webDriver.FindElement(By.Id("shareBtn"));

            shareBtn.Click();

            Thread.Sleep(2000);

            const string xPathMessage = "//*[@id=\"message\"]/span";

            IWebElement message = _webDriver.FindElement(By.XPath(xPathMessage));

            Assert.Equal(textMessage, message.GetAttribute("textContent"));

            _webDriver.Close();
        }

        [Fact]
        public void Message_RevokeAccessUser()
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string textMessage = "Доступ пользователя отозван.";

            const string loginUser = "one1";
            const string passwordUser = "111111";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            IWebElement filter = _webDriver.FindElement(By.Id("noteScopeFilter"));

            filter.SendKeys("Мои");

            Thread.Sleep(2000);

            const string xPathLi = "//*[@id=\"notesList\"]/li";
            IWebElement li = _webDriver.FindElement(By.XPath(xPathLi));

            li.Click();

            const string xPathBtn = "//*[@id=\"sharedUsersList\"]/li[1]/button";
            IWebElement revokeBtn = _webDriver.FindElement(By.XPath(xPathBtn));

            revokeBtn.Click();

            Thread.Sleep(2000);

            var alert_win = _webDriver.SwitchTo().Alert();

            alert_win.Accept();

            Thread.Sleep(2000);

            const string xPathMessage = "//*[@id=\"message\"]/span";

            IWebElement message = _webDriver.FindElement(By.XPath(xPathMessage));

            Assert.Equal(textMessage, message.GetAttribute("textContent"));

            _webDriver.Close();
        }

        [Fact]
        public void UsersChangeEntry()
        {
            _webDriver.Url = "https://test.webmx.ru/";

            const string loginUser = "one1";
            const string passwordUser = "111111";

            const string blockClass = "share-block";

            const string entry = "two";

            IWebElement inputLogin = _webDriver.FindElement(By.Id("authUsername"));
            IWebElement inputPassword = _webDriver.FindElement(By.Id("authPassword"));

            inputLogin.SendKeys(loginUser);
            inputPassword.SendKeys(passwordUser);

            IWebElement authBtn = _webDriver.FindElement(By.Id("authSubmit"));

            authBtn.Click();

            Thread.Sleep(2000);

            IWebElement searchInput = _webDriver.FindElement(By.Id("searchInput"));

            searchInput.SendKeys(entry);

            Thread.Sleep(2000);
            //*[@id="notesList"]/li[1]
            //*[@id="notesList"]/li[2]
            const string xPathLiOne = "//*[@id=\"notesList\"]/li[1]";
            IWebElement li = _webDriver.FindElement(By.XPath(xPathLiOne));

            li.Click();

            Thread.Sleep(2000);

            IWebElement deleteBtn = _webDriver.FindElement(By.Id("deleteBtn"));

            IWebElement block = _webDriver.FindElement(By.Id("shareBlock"));

            Assert.Equal(blockClass, block.GetAttribute("class"));

            Assert.True(deleteBtn.Enabled);

            const string xPathLiTwo = "//*[@id=\"notesList\"]/li[2]";
            IWebElement li2 = _webDriver.FindElement(By.XPath(xPathLiTwo));

            li2.Click();

            Thread.Sleep(2000);

            Assert.NotEqual(blockClass, block.GetAttribute("class"));

            Assert.False(deleteBtn.Enabled);

            _webDriver.Close();
        }
    }
}
