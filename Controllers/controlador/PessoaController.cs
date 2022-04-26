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

namespace ProjetoFechamentoAutomatico.Controllers
{

    public class PessoaController : Controller
    {



        public ActionResult Login(Pessoa pessoa)
        {
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
                if (login.ACESSO == "NAO")
                {
                    TempData["Pessoa"] = login.IDPESSOA;
                    return RedirectToAction(nameof(UpdateSenha));
                }
                TempData["Pessoa"] = login.IDPESSOA;

                ViewData["Exibir"] = null;
                return RedirectToAction(nameof(PessoaController.Home));

            }
            else
            {
                TempData["mensagemErro"] = "Usuario e/ou Senha Incorreto !!!";
                return RedirectToAction("Login", "Pessoa");
            }

        }

        [HttpGet]
        public ActionResult Home()
        {

            object idusuario = TempData.Peek("Pessoa");
            if (idusuario != null)

            {
                string idusuario1 = "" + idusuario;
                //exibe ou não tag cadastro de usuario dependdo do idUsuario logado
                if (idusuario1 == "1" || idusuario1 == "50017" || idusuario1 == "50018")
                {
                    ViewData["CadastroUsuario"] = "sim";
                }

                TempData["Usuario"] = idusuario;
                
                return View();
            }
            else
            {
                return RedirectToAction("Login", "Pessoa");
            }

        }


        public ActionResult Cadastro()
        {
            //pegando o idUsuario logado
            var idusuario = TempData["Usuario"];
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
                if (pessoa.IDPESSOA == 0 )
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
                    TempData["sucesso1"] = "Senha alterada com sucesso !!!";
                    return RedirectToAction(nameof(UpdateSenha));

                }
            }
            else
            {
                return RedirectToAction("Login", "Pessoa");
            }

        }
        [HttpPost]
        public IActionResult DDCVI(IFormFile formFile)
        {
            var idusuario = TempData["Usuario"];
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
            if(senha.USUARIO == null && senha.EMAIL == null)
            {
                TempData["Erro Senha 1"] = "Todos os campos devem ser preenchidos !!!";
                return RedirectToAction(nameof(EsqueciSenha));
            }
            ConexaoOracle conexaoOracle = new ConexaoOracle();
            var ser = conexaoOracle.EsqueciSenha(senha);
            if(ser[0].SENHA != null)
            {
                MailMessage mail = new MailMessage("informa.ti@valenca.com.br",senha.EMAIL);
                mail.Subject = "Recuperação de Senha";
                mail.Body = " Usuário: " + senha.USUARIO + " \n Senha: "+ ser[0].SENHA + "";
                mail.Sender = new MailAddress("informa.ti@valenca.com.br");

                SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
                smtp.EnableSsl = true;
                smtp.Credentials = new System.Net.NetworkCredential("informa.ti@valenca.com.br", "!@m07a880");
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
                mail.Subject = "Aprovação de Formulario DD-CVI de "+logado[0].NOME+"";
                mail.Body = "" + logado[0].NOME + " \n Setor: " + logado[0].SETOR + "\n Cargo: " + logado[0].CARGO + " \n enviou um link para aprovação de formulario DD-CVI  ";
                mail.Sender = new MailAddress("informa.ti@valenca.com.br");

                SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
                smtp.EnableSsl = true;
                smtp.Credentials = new System.Net.NetworkCredential("informa.ti@valenca.com.br", "!@m07a880");
                smtp.Send(mail);
                TempData["sucesso1"] = "Email Enviado com sucesso !!!";
                return RedirectToAction(nameof(EnviarEmail));
            }
            return View();

        }



    }


}


