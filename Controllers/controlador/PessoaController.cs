using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc.Rendering;
using ProjetoFechamentoAutomatico.Models.Pessoa;
using System.Security.Principal;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using System.Dynamic;
using ProjetoFechamentoAutomatico.Controllers.controlador;
using OfficeOpenXml;
using System.IO;
using AspNetCore.IServiceCollection;
using System.Net.Mail;
using System.Collections;
using System.Text;
using System.Security.Cryptography;
using System.Net;


namespace ProjetoFechamentoAutomatico.Controllers
{

    public class PessoaController : Controller
    {
      static object documento;
        public ActionResult Login(Pessoa pessoa)
        {
            if (TempData["Redirect"] != null)
            {
                TempData["sucesso1"] = "Senha alterada com sucesso !!!";
            }

            TempData.Remove("Pessoa");

            return View();
        }
        [HttpPost]
        public ActionResult Entrar(Pessoa pessoa)

        {
            ConexaoOracle log = new ConexaoOracle();
            var login = log.Conectar(pessoa);

            if (login.IDPESSOA != 0)
            {
                //redirect para alterar a senha caso seja o primeiro acesso do usuario
                if (login.ACESSO == "0")
                {
                    TempData["Pessoa"] = login.IDPESSOA;
                    return RedirectToAction(nameof(UpdateSenha));
                }
                //TempData["Pessoa"] = login.IDPESSOA;
                var id = Convert.ToString(login.IDPESSOA);
                var U = login.USUARIO;
                //criptografa o id pessoa para passar na url
                string d = null;
              
                using (MD5 md5 = MD5.Create())
                {
                    //criptografia id
                    byte[] inputBytes = Encoding.UTF8.GetBytes(id);
                    byte[] hash = md5.ComputeHash(inputBytes);
                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                    for (int i = 0; i < hash.Length; i++)
                    {
                        sb.Append(hash[i].ToString("x2"));
                    }                                    
                    d = sb.ToString();
                   
                }
                TempData["USUARIO"] = login.USUARIO;
                //TempData["Perfil"] = login.PERFIL;
                ViewData["Exibir"] = null;
                return RedirectToAction(nameof(PessoaController.Home), new { id}) ;

            }
            else
            {
                TempData["mensagemErro"] = "Usuario e/ou Senha Incorreto !!!";
                return RedirectToAction("Login", "Pessoa");
            }

        }

        [HttpGet]
        public ActionResult Home(Pessoa pessoa)
        {
            var idusuari = TempData.Peek("Pessoa");
            //string idcripto = Request.QueryString.Value.Split("?d=")[1];
            if (idusuari == null)
            {
                string idusuario = Request.Path.Value.Split("/Pessoa/Home/")[1];
                pessoa.IDPESSOA = Convert.ToInt32(idusuario);
            }
            else
            {

                pessoa.IDPESSOA = Convert.ToInt32(idusuari);
            }
            //Boolean compare = MD5.ReferenceEquals(idcripto, idusuario);
        
           
            //object USUARIO = TempData.Peek("USUARIO");

            ConexaoOracle log = new ConexaoOracle();
            var login = log.Conectarhome(pessoa);


            //object idusuario = TempData.Peek("Pessoa");
            //object perfil = TempData.Peek("Perfil");
            if (login.IDPESSOA > 0 )
            {
               
                string idusuario1 = "" + login.IDPESSOA;
                //exibe ou não tag cadastro de usuario dependdo do idUsuario logado
                if (idusuario1 == "1" || idusuario1 == "50017" || idusuario1 == "50018" || idusuario1 == "50019")
                {
                    ViewData["CadastroUsuario"] = "sim";
                }

                TempData["Pessoa"] = idusuario1;

            }
            else
            {
                return RedirectToAction("Login", "Pessoa");
            }
            // TempData["USUARIO"] = USUARIO;
            //  HttpWebRequest req = WebRequest.CreateHttp("https://localhost:44319/Pessoa/Home/");
            //req.Method = "GET";
            //req.Headers.Add("X-ICM-API-Authorization", idusuario);
            //req.UserAgent = "RequisicaoWebDemo";




            return View();
        }


        public ActionResult Cadastro()
        {
            //pegando o idUsuario logado
            var idusuario = TempData["Pessoa"];
            var redirect = TempData["redirect"];
            if (redirect != null)
            {
                TempData["Cadastro"] = "Todos os campos devem ser preenchidos !!!";
            }
            //permitindo entrar no cadastro de usuario somente com id especifico( cadastrar mais pessoas)
            string idusuario1 = "" + idusuario;
            if (idusuario1 == "1" || idusuario1 == "50017" || idusuario1 == "50018")
            {
                ViewData["Exibir"] = "Sim";
                ViewData["exibir login"] = "sim";
                TempData["Usuario"] = idusuario;
                return View();
            }

            else
            {
                return RedirectToAction(nameof(PessoaController.Home));
            }

        }


        public ActionResult SalvarUsario(Pessoa pessoa)
        {

            try
            {
                var idusuario = TempData["Usuario"];
                if (pessoa.PERFIL == "PERFIL")
                {
                    pessoa.PERFIL = "1";
                }
                else
                {
                    pessoa.PERFIL = "0";
                }
                if (pessoa.NOME == null || pessoa.USUARIO == null || pessoa.SETOR == null || pessoa.CARGO == null || pessoa.SENHA == null || pessoa.EMAIL == null || pessoa.PERFIL == null)
                {

                    TempData["Usuario"] = idusuario;
                    TempData["redirect"] = "sim";
                    return RedirectToAction(nameof(Cadastro));
                }
                ConexaoOracle conexaoOracle = new ConexaoOracle();
                conexaoOracle.inserirUsuario(pessoa);
                if (pessoa.IDPESSOA == 0)
                {
                    TempData["Erro Usuario"] = " Nome de usuário já Existente. Por favor cadastre o usário com outro nome !!!";
                    TempData["Usuario"] = idusuario;
                    return RedirectToAction(nameof(Cadastro));
                }
                else
                {
                    TempData["Usuario"] = idusuario;
                    return RedirectToAction(nameof(Cadastro));
                }


            }
            catch
            {

                return View("Não foi Possivel inserir usuário");

            }

        }

        public ActionResult UpdateSenha()
        {

            return View();
        }
        public ActionResult SalvarSenha(Pessoa pessoa)
        {
            int idpessoa = Convert.ToInt32(TempData["Pessoa"]);
            pessoa.IDPESSOA = idpessoa;
            if (pessoa.IDPESSOA != 0)
            {

                if (pessoa.SENHA != pessoa.CONFSENHA)
                {
                    TempData["mensagemErro"] = "As senhas não são iguais !!!";
                    return RedirectToAction(nameof(UpdateSenha));
                }
                else
                {
                    ConexaoOracle up = new ConexaoOracle();
                    var update = up.updateSenha(pessoa);
                    TempData["Redirect"] = "sim";

                    return RedirectToAction(nameof(Login));

                }
            }
            else
            {
                return RedirectToAction("Login", "Pessoa");
            }

        }
        public ActionResult DDCVI(Pessoa pessoa)
        {
            TempData.Remove("caminho");
            var idusuario = TempData["Pessoa"];
            var idusuario1 = Convert.ToInt32(idusuario);
            pessoa.IDPESSOA = idusuario1;
            ConexaoOracle ddcvi = new ConexaoOracle();
            var USUARIO1 = ddcvi.DDCVI(pessoa);
            TempData["USUARIO"] = USUARIO1[0].USUARIO;
            if (idusuario != null)
            {
                TempData["Pessoa"] = idusuario;
                TempData["USUARIO"] = USUARIO1[0].USUARIO;
                return View();

            }

            {
                return RedirectToAction(nameof(Home));
            }

            var USUARIO = TempData["USUARIO"];
        }
        [HttpPost]
        public IActionResult DDCVI(IFormFile formFile)
        {
            var idusuario = TempData["Pessoa"];
            var USUARIO = TempData["USUARIO"];
            if (idusuario != null)
            {
                if (formFile.Length > 0)
                {
                    string nome = formFile.FileName;
                    if (nome.Contains(".xlsx") || nome.Contains(".xls"))
                    {
                        try
                        {
                            //cria um caminho temporario o arquivo escolhido
                            var caminho = Path.GetTempFileName();
                            // salva o arquivo escolhido no caminho temporario criado
                            using (var stream = System.IO.File.Create(caminho))
                            {
                                formFile.CopyTo(stream);
                            }
                            //captura o caminho do arquivo temporario e salva em uma tempdata
                            TempData["caminho"] = caminho;
                            TempData["Usuario"] = idusuario;
                            TempData["USUARIO"] = USUARIO;
                            return RedirectToAction("DDCVI", "Formularios");

                        }
                        catch (Exception ex)
                        {
                            return View(ex.Message);
                        }
                    }
                    else
                    {
                        TempData["mensagemErro 1"] = "Escolha um arquivo excel valido";
                        return RedirectToAction(nameof(Home));
                    }

                }
                else
                {
                    TempData["mensagemErro 1"] = "Escolha um arquivo excel valido";
                    return RedirectToAction(nameof(Home));
                }
            }
            else
            {

                return RedirectToAction(nameof(Login));
            }

        }
        public IActionResult EsqueciSenha()
        {
            return View();
        }
        public IActionResult EnviarSenha(EsqueciSenha senha)
        {
            if (senha.USUARIO == null && senha.EMAIL == null)
            {
                TempData["Erro Senha 1"] = "Todos os campos devem ser preenchidos !!!";
                return RedirectToAction(nameof(EsqueciSenha));
            }
            ConexaoOracle conexaoOracle = new ConexaoOracle();
            var ser = conexaoOracle.EsqueciSenha(senha);
            if (ser[0].SENHA != null)
            {
                MailMessage mail = new MailMessage("informa.ti@valenca.com.br", senha.EMAIL);
                mail.Subject = "Recuperação de Senha";
                mail.IsBodyHtml = true;
                mail.Body = " Usuário: " + senha.USUARIO + " \n Senha: " + ser[0].SENHA + "";
                mail.Sender = new MailAddress("informa.ti@valenca.com.br");

                SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
                smtp.EnableSsl = true;
                smtp.Credentials = new System.Net.NetworkCredential("informa.ti@valenca.com.br", "******");
                smtp.Send(mail);
                TempData["sucesso"] = "Mensagem enviada com sucesso para email cadastrado";
                return RedirectToAction(nameof(EsqueciSenha));
            }
            else
            {
                TempData["Usuario"] = "Usuário e/ou Email invalidos !!!";
                return RedirectToAction(nameof(EsqueciSenha));
            }

        }
        [HttpGet]
        public ActionResult EnviarEmail()
        {
            ViewData["Exibir"] = "Sim";
            ConexaoOracle emails = new ConexaoOracle();
            var email = emails.Emails();
            ViewData["emails"] = new SelectList(email, "IDPESSOA", "EMAIL");
            return View();
        }
        [HttpPost]
        public ActionResult EnviarEmail(string emails)
        {
            var usuario = TempData["Usuario"];
            string pessoaUsuario = "" + usuario;
            ViewData["Exibir"] = "Sim";
            ConexaoOracle email = new ConexaoOracle();
            var emailenviado = email.Email(emails);
            if (emailenviado[0].EMAIL != null)
            {
                emails = pessoaUsuario;
                ConexaoOracle pessoa = new ConexaoOracle();
                var logado = pessoa.Email(emails);

                MailMessage mail = new MailMessage("informa.ti@valenca.com.br", emailenviado[0].EMAIL);
                mail.Subject = "Aprovação de Formulario DD-CVI de " + logado[0].NOME + "";
                mail.IsBodyHtml = true;
                mail.Body = "" + logado[0].NOME + " \n Setor: " + logado[0].SETOR + "\n Cargo: " + logado[0].CARGO + " \n enviou um link para aprovação de formulario DD-CVI.  \n " +
                    "\n Clique no link abaixo para entrar no sistema. \n  \n  https://localhost:44319/Pessoa/Login";
                mail.Sender = new MailAddress("informa.ti@valenca.com.br");

                SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
                smtp.EnableSsl = true;
                smtp.Credentials = new System.Net.NetworkCredential("informa.ti@valenca.com.br", "*******");
                smtp.Send(mail);
                TempData["sucesso1"] = "Email Enviado com sucesso !!!";
                return RedirectToAction(nameof(EnviarEmail));
            }
            return View();

        }
        public ActionResult TipoMovimento(Pessoa pessoa)
        {
            TempData.Remove("caminho");
            var idusuario = TempData["Pessoa"];
            var idusuario1 = Convert.ToInt32(idusuario);
            pessoa.IDPESSOA = idusuario1;
            ConexaoOracle ddcvi = new ConexaoOracle();
            var USUARIO1 = ddcvi.DDCVI(pessoa);
            TempData["USUARIO"] = USUARIO1[0].USUARIO;
            if (idusuario != null)
            {
                TempData["Pessoa"] = idusuario;
                TempData["USUARIO"] = USUARIO1[0].USUARIO;
                return View();

            }

            {
                return RedirectToAction(nameof(Home));
            }

            var USUARIO = TempData["USUARIO"];
        }
        [HttpPost]
        public IActionResult TipoMovimento(IFormFile formFile)
        {
            var idusuario = TempData["Pessoa"];
            var USUARIO = TempData["USUARIO"];
            if (idusuario != null)
            {
                if (formFile.Length > 0)
                {
                    string nome = formFile.FileName;
                    if (nome.Contains(".xlsx") || nome.Contains(".xls"))
                    {
                        try
                        {
                            //cria um caminho temporario o arquivo escolhido
                            var caminho = Path.GetTempFileName();
                            // salva o arquivo escolhido no caminho temporario criado
                            using (var stream = System.IO.File.Create(caminho))
                            {
                                formFile.CopyTo(stream);
                            }
                            //captura o caminho do arquivo temporario e salva em uma tempdata
                            TempData["caminho"] = caminho;
                            TempData["Usuario"] = idusuario;
                            TempData["USUARIO"] = USUARIO;
                            return RedirectToAction("TipoMovimento", "Formularios");

                        }
                        catch (Exception ex)
                        {
                            return View(ex.Message);
                        }
                    }
                    else
                    {
                        TempData["mensagemErro 1"] = "Escolha um arquivo excel valido";
                        return RedirectToAction(nameof(Home));
                    }

                }
                else
                {
                    TempData["mensagemErro 1"] = "Escolha um arquivo excel valido";
                    return RedirectToAction(nameof(Home));
                }

            }
            else
            {

                return RedirectToAction(nameof(Login));
            }

        }
        public ActionResult TipoMovimentoContabil(Pessoa pessoa)
        {
            var idusuari = TempData["Pessoa"];        
            if(idusuari == null)
            {
                var teste = ((Microsoft.Extensions.Primitives.StringValues[])Request.Headers.Values)[6];
                var idusuario = teste.ToString().Split("/Pessoa/Home/")[1];
                var idusuario1 = Convert.ToInt32(idusuario);
                pessoa.IDPESSOA = idusuario1;
            }
            else
            {
                var idusuario = idusuari;                
                var idusuario1 = Convert.ToInt32(idusuario);
                pessoa.IDPESSOA = idusuario1;
            }
            //var idusuario = teste.ToString().Split("/Pessoa/Home/")[1];           
            //var idusuario1 = Convert.ToInt32(idusuario);
            //pessoa.IDPESSOA = idusuario1;
            ConexaoOracle ddcvi = new ConexaoOracle();
            var USUARIO1 = ddcvi.DDCVI(pessoa);          
            if (pessoa.IDPESSOA != 0)
            {
                var idusuario = Convert.ToString(pessoa.IDPESSOA);
                //TempData["Pessoa"] = idusuario;
               var usuario = USUARIO1[0].USUARIO;
                HttpWebRequest req = WebRequest.CreateHttp("https://localhost:44319/Pessoa/TipoMovimentoContabil/");             
                req.Method = "GET";
                req.Headers.Add(idusuario, "idusuario");
                req.Headers.Add(usuario, "usuario");
                documento = req;
                TempData["Pessoa"] = idusuario;
               
                return View();

            }

            {
                return RedirectToAction(nameof(Home));
            }

            var USUARIO = TempData["USUARIO"];
            
        }
        [HttpPost]
        public IActionResult TipoMovimentoContabil(IFormFile formFile)
        {
            var doc = documento;
            var id = ((System.Net.HttpWebRequest)doc).Headers.AllKeys[0];
      
            
            if (id != null)
            {
                if (formFile.Length > 0)
                {
                    string nome = formFile.FileName;
                    if (nome.Contains(".xlsx") || nome.Contains(".xls"))
                    {
                        try
                        {
                            //cria um caminho temporario o arquivo escolhido
                            var caminho = Path.GetTempFileName();
                            // salva o arquivo escolhido no caminho temporario criado
                            using (var stream = System.IO.File.Create(caminho))
                            {
                                formFile.CopyTo(stream);
                            }
                           
                            return RedirectToAction("TipoMovimentoContabil", "Formularios", new {id, caminho});
                        }
                        catch (Exception ex)
                        {
                            return View(ex.Message);
                        }
                    }
                    else
                    {
                        TempData["mensagemErro 1"] = "Escolha um arquivo excel valido";
                        return RedirectToAction(nameof(Home));
                    }

                }
                else
                {
                    TempData["mensagemErro 1"] = "Escolha um arquivo excel valido";
                    return RedirectToAction(nameof(Home));
                }

            }
            else
            {

                return RedirectToAction(nameof(Login));
            }

        }

    }
}





