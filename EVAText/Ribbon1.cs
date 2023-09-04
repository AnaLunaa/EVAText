using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

// TODO:  Siga estos pasos para habilitar el elemento (XML) de la cinta de opciones:

// 1: Copie el siguiente bloque de código en la clase ThisAddin, ThisWorkbook o ThisDocument.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Cree métodos de devolución de llamada en el área "Devolución de llamadas de la cinta de opciones" de esta clase para controlar acciones del usuario,
//    como hacer clic en un botón. Nota: si ha exportado esta cinta de opciones desde el diseñador de la cinta de opciones,
//    mueva el código de los controladores de eventos a los métodos de devolución de llamada y modifique el código para que funcione con el
//    modelo de programación de extensibilidad de la cinta de opciones (RibbonX).

// 3. Asigne atributos a las etiquetas de control del archivo XML de la cinta de opciones para identificar los métodos de devolución de llamada apropiados en el código.  

// Para obtener más información, vea la documentación XML de la cinta de opciones en la Ayuda de Visual Studio Tools para Office.


namespace EVAText
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        #region Miembros de IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("EVAText.Ribbon1.xml");
        }

        #endregion

        #region Devoluciones de llamada de la cinta de opciones
        //Cree métodos de devolución de llamada aquí. Para obtener más información sobre la adición de métodos de devolución de llamada, visite https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }
        public void Micro1(Office.IRibbonControl control)
        {
            //Concordancia de número

            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Debes tener en cuenta sí el elemento al que te refieres está en singular o plural.");
        }
        public void Micro2(Office.IRibbonControl control)
        {
            //Concordancia de género
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Debes tener en cuenta sí el elemento al que te refieres está en femenino o masculino.");
        }
        public void Micro3(Office.IRibbonControl control)
        {
            //Repetición léxica
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Puedes evitar la repetición léxica usando sinónimos o pronombres.");
        }
        public void Micro4(Office.IRibbonControl control)
        {
            //Imprecisión por homonimia
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Revisa la homonimia de las palabras, estas confundiendo haya con halla, ahí con hay, vaya con valla...");
        }
        public void Micro5(Office.IRibbonControl control)
        {
            //Imprecisión por polisemia
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Considera usar sinónimos en lugar de palabras parónimas (una palabra con múltiples significados).");
        }
        public void Micro6(Office.IRibbonControl control)
        {
            //Queísmo
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "En este caso debes usar la preposición de antes de la conjunción que.");
        }
        public void Micro7(Office.IRibbonControl control)
        {
            //Dequeísmo
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "En este caso debes omitir la preposición de antes de la conjunción que.");
        }
        public void Micro8(Office.IRibbonControl control)
        {
            //Ausencia de predicado
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Qué es lo que pasa con esta acción. Es necesario para que se entienda la cualidad, propiedad o estado del sujeto o del complemento directo que estas mencionando.");
        }
        public void Micro9(Office.IRibbonControl control)
        {
            //Tiempo verbal incongruente
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Es importante mantener la consistencia en el uso de los tiempos verbales dentro de un párrafo. Te sugiero revisar y asegurarte de que los verbos se ajusten adecuadamente al contexto y tiempo narrativo que deseas transmitir.");
        }
        public void Micro10(Office.IRibbonControl control)
        {
            //Modo verbal incongruente
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Asegúrate de que el modo verbal seleccionado se ajuste al sentido que deseas transmitir, si es una orden emplea el imperativo, si es un deseo el public voidjuntivo o si es una afirmación el indicativo.");
        }
        public void Micro11(Office.IRibbonControl control)
        {
            //Condicional incongruente
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "El condicional se utiliza para expresar situaciones hipotéticas o posibles y sus resultados. Revisa las oraciones donde empleas esta estructura y asegúrate de que estén correctamente construidas.");
        }
        public void MD1(Office.IRibbonControl control)
        {
            //MD1 comentadores
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para presentar un nuevo comentario, puedes usar: pues, pues bien, dicho esto/eso, así las cosas.");
        }
        public void MD2(Office.IRibbonControl control)
        {
            //MD2 ordenadores
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para agrupan varios elementos como partes de un solo comentario, puedes usar: en primer lugar/en segundo lugar; por una parte/por otra parte; de un lado/de otro lado.");
        }
        public void MD3(Office.IRibbonControl control)
        {
            //MD3 digresores
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para introducir un comentario en relación con el tópico principal del discurso, puedes usar: por cierto, a propósito, a todo esto.");
        }

        public void MD4(Office.IRibbonControl control)
        {
            //MD4 aditivos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para  unir un elemento con otro de la misma orientación argumentativa, puedes usar:  además, encima, aparte, incluso.");
        }
        public void MD5(Office.IRibbonControl control)
        {
            //MD5 consecutivos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para vincular un miembro de discurso con otro previo o con una suposición contextual, puedes usar: por tanto, por consiguiente, por ende, en consecuencia, de ahí, entonces, pues, así, así pues.");
        }
        public void MD6(Office.IRibbonControl control)
        {
            //MD6 contraargumentativos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para unir dos miembros de manera que el segundo sea supresor o atenuador de alguna conclusión a la se pudiera obtener del primero, puedes usar: en cambio, por el contrario, por contra, antes bien, sin embargo, no obstante, con todo...");
        }
        public void MD7(Office.IRibbonControl control)
        {
            //MD7 explicativos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para introducir una nueva formulación de lo que se ha enunciado en el discurso previo, puedes usar: o sea, es decir, esto es, a saber; en otras palabras.");
        }
        public void MD8(Office.IRibbonControl control)
        {
            //MD8 de rectificación
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para sustituir un primer miembro que presentas como una formulación incorrecta, por otra que la corrige, o al menos la mejora, puedes usar: mejor dicho, mejor aún, más bien.");
        }
        public void MD9(Office.IRibbonControl control)
        {
            //MD9 de distanciamiento
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para presentar como no relevante un miembro del discurso ya mencionado, puedes usar en cualquier caso, en todo caso, de todos modos.");
        }
        public void MD10(Office.IRibbonControl control)
        {
            //MD10 recapitulativos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para presentar un elemento del discurso como una conclusión o recapitulación a partir de un elemento anterior, puedes usar: en suma, en conclusión, en definitiva, en fin, al fin y al cabo.");
        }
        public void MD11(Office.IRibbonControl control)
        {
            //MD11 de refuerzo argumentativo
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para reforzar como argumento un elemento del discurso, puedes usar: en realidad, en el fondo, de hecho...");
        }
        public void MD12(Office.IRibbonControl control)
        {
            //MD12 de concreción
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para concretar o mostrar un ejemplo de una expresión más general, puedes usar: por ejemplo, en particular.");
        }
        public void MD13(Office.IRibbonControl control)
        {
            //MD13 epistémicos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "En enunciados declarativos, puedes usar marcadores discursivos como: claro, desde luego, por supuesto, naturalmente y sin duda,  en efecto, por lo visto, al parecer..");
        }
        public void Cohesion1(Office.IRibbonControl control)
        {
            //Cohesion Aditiva
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para unir o enlazar dos o más componentes de una oración, puedes usar las conjunciones y, e, ni, que.");
        }
        public void Cohesion2(Office.IRibbonControl control)
        {
            //Cohesion Adversativa
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para presentar oposición o diferencia, entre la frase anterior y la que sigue, puedes usar las conjunciones pero, mas, empero, sino, aunque, sin embargo, no obstante, antes, antes bien, por lo demás.");
        }
        public void Cohesion3(Office.IRibbonControl control)
        {
            //Cohesion Causal
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para expresar una causa, motivo o razón, puedes usar las conjunciones porque, como, dado que, visto que, puesto que, pues, ya que.");
        }
        public void Cohesion4(Office.IRibbonControl control)
        {
            //Cohesion Temporal
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para expresar temporalidad, puedes usar las conjunciones mientras, mientras que, cuando, antes que, después que, aún no, luego que.");
        }
        public void Cohesion5(Office.IRibbonControl control)
        {
            //Repetición léxica
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Puedes evitar la repetición léxica usando sinónimos o pronombres.");
        }
        public void Cohesion6(Office.IRibbonControl control)
        {
            //repetición semántica
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Usa términos que están relacionados por su significado. Pueden ser sinónimos, antónimos, hiperónimos o hipónimos.");
        }
        public void Cohesion7(Office.IRibbonControl control)
        {
            //Repetición sintáctica
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Hay una repetición de esquemas sintácticos, es decir, oraciones con el mismo orden.");
        }
        public void Cohesion8(Office.IRibbonControl control)
        {
            //Repetición anafórica
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "La recurrencia anafórica te permite  mantener el hilo del texto a partir de la utilización de pronombres y algunos adverbios.");
        }
        public void Cohesion9(Office.IRibbonControl control)
        {
            //Sinonimia
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Puedes evitar la reiteración excesiva de una determinada palabra haciendo uso de sinónimos, hiperónimos e hipónimos.");
        }
        public void Cohesion10(Office.IRibbonControl control)
        {
            //Referencia pronominal
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Debes aclarar el atributo. Para ello, emplea los pronombres de complemento: la, le o lo o los posesivos: a mi, a ti, a él...");
        }
        public void Cohesion11(Office.IRibbonControl control)
        {
            //Sujeto discursivo
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Ten en cuenta que el narrador discursivo es académico y formal. Lo mejor es usar la tercera persona del singular si es un trabajo individual y si es uno grupal, emplea la primera persona del plural. Esta debe concordar a lo largo del texto.");
        }
        public void Cohesion12(Office.IRibbonControl control)
        {
            //Referencia personal
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Debes usar el pronombres personales o posesivo correcto: mí, tu, su, mío, tuyo, suyo...");
        }
        public void Cohesion13(Office.IRibbonControl control)
        {
            //Referencia temporal
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "No queda claro el tiempo al que te refieres. Emplea deícticos temporales. antes, ahora, después, ayer, hoy, mañana...");
        }
        public void Cohesion14(Office.IRibbonControl control)
        {
            //Referencia espacial
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "No queda claro el espacio al que te refieres. Emplea deícticos espaciales: aquí, acá; ahí, arriba, abajo, encima...");
        }
        public void Cohesion15(Office.IRibbonControl control)
        {
            //Referencia comparativa
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que una comparación tiene dos elementos y usa estructuras como: tan(to/a/os/as.)...como, igual... que, más/menos..que...");
        }
        public void Coherencia1(Office.IRibbonControl control)
        {
            //Enunciador public voidjetivo
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Procura aplicar las estrategias de despersonalización (voz pasiva con ‘ser’, formas impersonales y pasivas con ‘se’, nominalizaciones), evita el uso de la primera persona en singular.");
        }
        public void Coherencia2(Office.IRibbonControl control)
        {
            //Muletillas
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Revisa y evita las muletillas, pueden resultar expresiones innecesarias y repetitivas.");
        }
        public void Coherencia3(Office.IRibbonControl control)
        {
            //Verbos valorativos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Evita la public voidjetividad sin hacer uso de verbos valorativos como creer, sentir, opinar.");
        }
        public void Coherencia4(Office.IRibbonControl control)
        {
            //léxico coloquial
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "En el lenguaje académico evita usar términos coloquiales o emplear un registro particular del lenguaje oral, esto le resta formalidad al texto.");
        }
        public void Coherencia5(Office.IRibbonControl control)
        {
            //Verbo haber
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que el verbo haber solo se conjuga en género y número cuando es verbo principal. Pero cuando es una estructura impersonal siempre va en singular ya que no hay sujeto gramatical.");
        }
        public void Coherencia6(Office.IRibbonControl control)
        {
            //Uso de public voidjuntivo
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Presta atención a los verbos en contextos de incertidumbre, deseo, opinión o posibilidad. Asegúrate de emplear correctamente el public voidjuntivo en estas situaciones");
        }
        public void Coherencia7(Office.IRibbonControl control)
        {
            //Contradicciones
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Esta idea que mencionas se contradice con lo que dijiste anteriormente.");
        }
        public void Coherencia8(Office.IRibbonControl control)
        {
            //Redundancias
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Esto que mencionas es innecesario porque ya lo dijiste anteriormente.");
        }
        public void Coherencia9(Office.IRibbonControl control)
        {
            //Perdida del tema
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "No queda clara la relación de esta idea con el tema y propósito del escrito. Parece que estuvieras hablando de otra cosa diferente.");
        }
        public void Coherencia10(Office.IRibbonControl control)
        {
            //Salto temáticos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Es importante que cuando abordes un public voidtema digas todo sobre él y luego pases al siguiente no que vayas brincando entre todos los public voidtemas porque será muy difícil la comprensión.");
        }
        public void Coherencia11(Office.IRibbonControl control)
        {
            //Incoherencia
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "No queda clara la relación de esta idea con la anterior y/o la siguiente.");
        }
        public void Coherencia12(Office.IRibbonControl control)
        {
            //Ambigüedad temática
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Falta trabajar la claridad y la precisión, esta idea puede tener varios significados.");
        }
        public void Estructura1(Office.IRibbonControl control)
        {
            //Extensión del título
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Procura ser claro y preciso sin hacer uso de demasiadas palabras. (Máximo 12).");
        }
        public void Estructura2(Office.IRibbonControl control)
        {
            //Estructura del título
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Ten cuidado con la sintaxis y la precisión léxica del título, procura tener un orden adecuado de las palabras y evita la polisemia. También evita el uso de abreviaturas, paréntesis o caracteres desconocidos.");
        }
        public void Estructura3(Office.IRibbonControl control)
        {
            //Retórica textual del título
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Trata en lo posible de indicar una acción en el título, además procura que sea denotativo o informativo.");
        }
        public void Estructura4(Office.IRibbonControl control)
        {
            //Tema del título
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que el titulo debe ser coherente con el contenido y las ideas principales del escrito.");
        }
        public void Estructura5(Office.IRibbonControl control)
        {
            //Objetivos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Es importante que el punto de vista, objetivo del trabajo y sus intenciones queden claras desde el inicio.");
        }
        public void Estructura6(Office.IRibbonControl control)
        {
            //Planteamiento del problema
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Es necesario aclarar el problema y el contexto del tema.");
        }
        public void Estructura7(Office.IRibbonControl control)
        {
            //Antecedentes
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Antes de que pases al contenido, aclara los antecedentes del tema y los principales referentes teóricos que consultaste.");
        }
        public void Estructura8(Office.IRibbonControl control)
        {
            //Enumeración del contenido
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Es bueno que le aclares al lector los public voidtemas que desarrollaras en tu trabajo.");
        }
        public void Estructura9(Office.IRibbonControl control)
        {
            //Soluciones
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda mencionar las posibles soluciones al problema que mencionaste en la introducción y abordaste en el contenido.");
        }
        public void Estructura10(Office.IRibbonControl control)
        {
            //Aplicaciones
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Sería importante que mencionaras las aplicaciones prácticas del tema que estas abordando.");
        }
        public void Estructura11(Office.IRibbonControl control)
        {
            //Hipótesis
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda explicar de manera clara y precisa cómo se cumplieron los objetivos mencionados en la introducción y abordados en el contenido o si el punto de vista fue valido y{o hasta qué punto.");
        }
        public void Estructura12(Office.IRibbonControl control)
        {
            //Implicaciones
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Una buena de resumir el escrito en las conclusiones es enumerando los pro y los contras del tema y los principales public voidtemas abordados.");
        }

        public void citacion1(Office.IRibbonControl control)
        {
            //Cita textual
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "La cita textual es una reproducción literal no mayor a 40 palabras, va entre comillas y debes incluir la página.");
        }
        public void citacion2(Office.IRibbonControl control)
        {
            //Cita narrativa o parafraseada
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "En la cita parafraseada se dice con palabras propias las ideas de otra fuente. No va entre comillas, ni lleva número de página.");
        }
        public void citacion3(Office.IRibbonControl control)
        {
            //Cita en bloque
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "La cita en bloque es una reproducción literal mayor a 40 palabras. No va entre comillas pero si en un párrafo aparte y debes incluir el número de página.");
        }
        public void citacion4(Office.IRibbonControl control)
        {
            //Referencia de la cita en la lista
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Este texto no está enlistado en las referencias empleadas para hacer el trabajo.");
        }
        public void citacion5(Office.IRibbonControl control)
        {
            //Autenticidad
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Es imprescindible que tu texto sea original, que sean tus ideas y tus palabras. Pero asegúrate de incluir referencias precisas cada vez que utilices ideas, datos o palabras de otros autores.");
        }
        public void citacion6(Office.IRibbonControl control)
        {
            //Cita directa
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Cuando reproduces las ideas de otro debes aclarar quién es. Para ello es importante que escribas su apellido y el año de la publicación entre el paréntesis (En APA) o escribas el superíndice (en IEEE) al final.");
        }
        public void citacion7(Office.IRibbonControl control)
        {
            //Cita indirecta
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Cuando dentro de la idea que citas empleas el apellido del autor debes aclarar el año de la publicación entre paréntesis (en APA) o el superíndice (en IEEE).");
        }
        public void citacion8(Office.IRibbonControl control)
        {
            //Lista de referencias
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Este texto no hace parte del contenido.");
        }
        public void citacion9(Office.IRibbonControl control)
        {
            //Datos de referencias
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Teniendo en cuenta el tipo de texto que citaste, debes revisar cómo se organizaran los elementos: los datos del autor, el título, la editorial...");
        }
        public void citacion10(Office.IRibbonControl control)
        {
            //Cita de definición
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que cuando estes definiendo conceptos debes decir de dónde sale.");
        }
        public void citacion11(Office.IRibbonControl control)
        {
            //Cita confirmatoria
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para corroborar/respaldar esta afirmación debes citar.");
        }
        public void citacion12(Office.IRibbonControl control)
        {
            // Cita posicionamiento
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "En la cita de posicionamiento el autor toma posición con respecto a la fuente que cita, desde una perspectiva crítica.");
        }
        public void citacion13(Office.IRibbonControl control)
        {
            //Cita dialéctica
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "En la cita dialéctica el autor compara a varios autores para establecer comparaciones o aproximaciones.");
        }
        public void citacion14(Office.IRibbonControl control)
        {
            //Cita de apoyo o expansión
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Crea una cita de apoyo  para indicarle al lector que existe información adicional para ser consultada. Generalmente se utiliza el término véanse.");
        }
        public void citacion15(Office.IRibbonControl control)
        {
            //Títulos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que el título principal debe ir centrado, en negrita y con cada letra inicial en mayúscula.");
        }
        public void citacion16(Office.IRibbonControl control)
        {
            //títulos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que el título debe ir alineado a la izquierda, en negrita y con cada letra inicial en mayúscula.");
        }
        public void citacion17(Office.IRibbonControl control)
        {
            //Notas al pie
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Las notas al pie se usan para expandir información de los autores del texto o la información contenida en él, no para dar los datos de una referencia.");
        }

        public void citacion18(Office.IRibbonControl control)
        {
            //Ilustraciones y gráficas
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que las tablas e ilustraciones deben estar enumeradas, seguidas de su título y debajo una nota aclaratoria con las fuentes de información empleadas.");
        }
        public void citacion19(Office.IRibbonControl control)
        {
            //Índices o tablas de contenido
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Cuando se escriben textos muy largos es importante crear un índice donde queden claros los public voidtítulos y las páginas en las que el lector puede encontrarlos.");
        }
        public void ED1(Office.IRibbonControl control)
        {
            //Espaciado
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que los trabajos académicos se emplea el espaciado 1.5 para facilitar la lectura y la posterior corrección.");
        }
        public void ED2(Office.IRibbonControl control)
        {
            //Letra
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que los trabajos académicos se emplea el espaciado 1.5 para facilitar la lectura y la posterior corrección.");
        }
        public void ED3(Office.IRibbonControl control)
        {
            //Signos de exclamación
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que las expresiones que denotan  sorpresa, asombro, alegría, súplica, mandato, deseo.. van entre signos de exclamación.");
        }
        public void ED4(Office.IRibbonControl control)
        {
            //Signos de interrogación
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda emplear signos de interrogación para delimitar las expresiones interrogativas y exclamativas directas.");
        }
        public void ED5(Office.IRibbonControl control)
        {
            //Puntos suspensivos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que los puntos suspensivos pueden usarse al final de enumeraciones abiertas o incompletas, para reproducir una cita textual omitiendo una parte del final o para expresar duda. Van dentro de paréntesis (...) o corchetes [...] cuando al transcribir literalmente un texto se omite una parte de él.");
        }
        public void ED6(Office.IRibbonControl control)
        {
            //Sangría
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que para facilitar la comprensión de los párrafos se emplea sangría de primera línea en los párrafos sencillos, mientras que en las citas en bloque, los encabezados y las tablas y figuras no se usa.");
        }
        public void ED7(Office.IRibbonControl control)
        {
            //Sesgos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Es importante que respetes el lenguaje empleado por las personas para describirse a sí mismas; que escojas etiquetas con sensibilidad con las que te puedas asegurar de respetar la individualidad y la humanidad de las personas y evita las falsas jerarquías.");
        }
        public void ED8(Office.IRibbonControl control)
        {
            // Macro reglas textuales
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, " Para exponer la información relevante de un texto, puedes hacer uso de las siguientes macrorreglas: supresión (omitir información innecesaria), selección (elegir lo relevante), generalización (abstraer características comunes) e integración (fusionar conceptos.");
        }
        public void ED9(Office.IRibbonControl control)
        {
            //Atenuación con usted
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que en los textos formales se emplea el usted ");
        }
        public void ED10(Office.IRibbonControl control)
        {
            //Atenuación del acto amenazador
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda emplear expresiones rituales como por favor, muchas gracias...");
        }
        public void ED11(Office.IRibbonControl control)
        {
            // Atenuación con se
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda emplear se para minimizar la amenaza a la propia imagen del hablante, para atenuar la fuerza de lo dicho, para eludir la responsabilidad");
        }
        public void ED12(Office.IRibbonControl control)
        {
            //Condicional de cortesía
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para expresar solicitud, sugerencia, ruego es mejor emplear el condicional: ¿Podrías abrir la ventana?");
        }
        public void ED13(Office.IRibbonControl control)
        {
            //Condicional de modestia
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para expresar una postura que puede estar en contradicción con la de los interlocutores es mejor emplear el condicional: diría que eso no es así.");
        }
        public void ED14(Office.IRibbonControl control)
        {
            //Plural inclusivo, sociativo o de modestia
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Para demostrar complicidad, un valor universal se emplea el plural o para mostrar modestia.");
        }
        public void Delimitadores1(Office.IRibbonControl control)
        {
            //Punto seguido
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que para separar enunciados que integran un mismo párrafo se debe usar el punto seguido.");
        }
        public void Delimitadores2(Office.IRibbonControl control)
        {
            //Punto aparte
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que para separar dos párrafos distintos, que suelen desarrollar, dentro de la unidad del texto,contenidos diferentes, debe ir un punto aparte.");
        }
        public void Delimitadores3(Office.IRibbonControl control)
        {
            //Punto y coma de separación(proposiciones yuxtapuestas)
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que se usa punto y coma para separar proposiciones yuxtapuestas, especialmente cuando en estas se ha empleado la coma.");
        }
        public void Delimitadores4(Office.IRibbonControl control)
        {
            //Punto y coma de enumeración compleja
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que se usa punto y coma para separar los elementos de una enumeración cuando se trata de expresiones complejas que incluyen comas.");
        }
        public void Delimitadores5(Office.IRibbonControl control)
        {
            //Punto y coma de conjunción
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que se usa punto y coma delante de las conjunciones adversativas pero, más y aunque cuando las oraciones que encabezan tienen cierta longitud.");
        }

        public void Delimitadores6(Office.IRibbonControl control)
        {
            //Dos puntos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que los dos puntos sirven para detener el discurso para llamar la atención sobre lo que sigue, se usan para anunciar enumeración, antes de citas textuales, para conectar oraciones y para presentar una conclusión o explicación.");
        }
        public void Delimitadores7(Office.IRibbonControl control)
        {
            //Coma enumerativa
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "La coma enumerativa se usa separar un conjunto de elementos que expresan características similares o simplemente se desean enumerar. Al nombrar estas palabras se usan conjunciones (y, o, u, ni), antes de ellas no se debe colocar una coma. ");
        }
        public void Delimitadores8(Office.IRibbonControl control)
        {
            //Coma vocativa
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "La coma vocativa se usa para marcar una diferencia entre el vocativo y el resto de la oración. El vocativo es la forma de dirigirse a una persona o más por su nombre o algo que lo distinga.");
        }
        public void Delimitadores9(Office.IRibbonControl control)
        {
            //Coma elíptica
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "La coma elíptica es utilizada para evitar redundancia en las oraciones. Se puede usar para sustituir un verbo o sustantivo que fue mencionado.");
        }
        public void Delimitadores10(Office.IRibbonControl control)
        {
            //Coma explicativa
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "La coma explicativa o incidental es utilizada para agregar datos adicionales del sujeto o el verbo.");
        }
        public void Delimitadores11(Office.IRibbonControl control)
        {
            //Coma apositiva
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, " La explicación que va entre las comas apositivas es un nombre o grupo de nombres. Ej.: María, mi vecina, regó las plantas en mis vacaciones.");
        }
        public void Delimitadores12(Office.IRibbonControl control)
        {
            // Coma hiperbática
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que cuando se altera el orden usual que debe tener una oración en cuanto al sujeto, el verbo y la acción, se debe usar la coma hiperbática.");
        }
        public void Delimitadores13(Office.IRibbonControl control)
        {
            //Coma conjuntiva
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que cuando en la oración realizas una pausa con alguna frase adverbial o conjunciones, debes usar la coma conjuntiva.");
        }
        public void Delimitadores14(Office.IRibbonControl control)
        {
            //Guion
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Cuando se escriben textos muy largos es importante crear un índice donde queden claros los public voidtítulos y las páginas en las que el lector puede encontrarlos.");
        }
        public void Delimitadores15(Office.IRibbonControl control)
        {
            //Parentesis
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Cuando se emplea el discurso directo se emplean los parentesis para aclarar el apellido, la fecha de publicación de la cita y el numero de pagina si es textual. Mientras que si se emplea el discurso indirecto en el parentesis solo va la fecha de publicación y el numero de pagina en el caso que sea textual.");
        }
        public void Delimitadores16(Office.IRibbonControl control)
        {
            //Comillas
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Cuando se realiza una cita textual de menos de 40 palabras se debe poner entre comillas.");
        }
        #endregion

        #region Asistentes

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
