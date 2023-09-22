using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
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
            currentRange.Comments.Add(currentRange, "Debes tener en cuenta si el elemento al que te refieres está en singular o plural. ");
        }
        public void Micro2(Office.IRibbonControl control)
        {
            //Concordancia de género
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Debes tener en cuenta si el elemento al que te refieres está en femenino o masculino.");
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
            currentRange.Comments.Add(currentRange, "Revisa la homonimia de las palabras, estás confundiendo haya con halla, ahí con hay, vaya con valla...");
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
            currentRange.Comments.Add(currentRange, "Debes agregar qué es lo que pasa con esta acción. Es necesario para que se entienda la cualidad, propiedad o estado del sujeto o del complemento directo que estás mencionando. ");
        }
        public void Micro9(Office.IRibbonControl control)
        {
            //Tiempo verbal incongruente
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Es importante que mantengas la consistencia en el uso de los tiempos verbales dentro de un párrafo. Te sugiero revisar y asegurarte de que los verbos se ajusten adecuadamente al contexto y tiempo narrativo que deseas transmitir.");
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
            currentRange.Comments.Add(currentRange, "embros de manera que el segundo sea supresor o atenuador de alguna conclusión a la se pudiera obtener del primero, puedes usar: en cambio, por el contrario, por contra, antes bien, sin embargo, no obstante, con todo...");
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
            //Enunciador subjetivo
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
            //Uso del subjuntivo
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Presta atención a los verbos en contextos de incertidumbre, deseo, opinión o posibilidad. Asegúrate de emplear correctamente el subjuntivo en estas situaciones. ");
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
            currentRange.Comments.Add(currentRange, " Parece que estuvieras hablando de otra cosa diferente. No queda clara la relación de esta idea con el tema y propósito del escrito.");
        }
        public void Coherencia10(Office.IRibbonControl control)
        {
            //Salto temáticos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Es importante que cuando abordes un subtema digas todo sobre él y luego pases al siguiente, no que presentes todos los subtemas al mismo tiempo porque será difícil la comprensión.");
        }
        public void Coherencia11(Office.IRibbonControl control)
        {
            //Incoherencia
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "No queda clara la relación de esta idea con la anterior y/o la siguiente. Procura Procura especificar mejor.");
        }
        public void Coherencia12(Office.IRibbonControl control)
        {
            //Ambigüedad temática
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Falta trabajar la claridad y la precisión, esta idea puede tener varios significados.");
        }
        public void Coherencia13(Office.IRibbonControl control)
        {
            //Desequilibrio
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Estás combinando párrafos largos con cortos sin una razón evidente. Es importante mantener un equilibrio. Recuerda que cada párrafo debe estar relacionado con una idea principal y al menos una o cinco ideas secundarias que la respalden.");
        }
        public void Coherencia14(Office.IRibbonControl control)
        {
            //Repetición
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "En este párrafo estás repitiendo algo que ya mencionaste. Te sugiero revisar la idea que quieres trasmitir y reescribirla. ");
        }
        public void Coherencia15(Office.IRibbonControl control)
        {
            //Párrafo frase
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Ten cuidado, ya que un párrafo debe tener una idea principal y, como mínimo, una secundaria, pero aquí solo has escrito una frase. Refuerza  lo que quieres trasmitir con una o dos ideas secundarias.");
        }
        public void Coherencia16(Office.IRibbonControl control)
        {
            //Párrafo extenso
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Debido a la longitud de este párrafo, considera dividirlo en párrafos más cortos para mejorar la claridad y hacer la lectura más fácil. Cada uno con su propia idea central. ");
        }
        public void Coherencia17(Office.IRibbonControl control)
        {
            //Inconcluso
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Esta idea parece estar incompleta. Te recomiendo revisar su redacción para que se comprenda desde el principio hasta el final.");
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
            currentRange.Comments.Add(currentRange, "Es necesario que aclares el problema y el contexto del tema.");
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
            currentRange.Comments.Add(currentRange, "Es bueno que le aclares al lector los subtemas que desarrollarás en tu trabajo. Así como las principales fuentes de información que empleaste y para qué.");
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
            currentRange.Comments.Add(currentRange, "Recuerda explicar de manera clara y precisa cómo se cumplieron los objetivos mencionados en la introducción y abordados en el contenido, o si el punto de vista fue válido y//o hasta qué punto.");
        }
        public void Estructura12(Office.IRibbonControl control)
        {
            //Implicaciones
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Una buena estrategia de resumir el escrito en las conclusiones es enumerando los pro y los contras del tema y los principales subtemas abordados. Puedes aplicarlo.");
        }

        public void citacion1(Office.IRibbonControl control)
        {
            //Cita textual
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que la cita textual es una reproducción literal no mayor a 40 palabras, va entre comillas y debes incluir la página.");
        }
        public void citacion2(Office.IRibbonControl control)
        {
            //Cita narrativa o parafraseada
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que en la cita parafraseada se dice con palabras propias las ideas de otra fuente. No va entre comillas, ni lleva número de página.");
        }
        public void citacion3(Office.IRibbonControl control)
        {
            //Cita en bloque
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que la cita en bloque es una reproducción literal mayor a 40 palabras. No va entre comillas pero si en un párrafo aparte y debes incluir el número de página.");
        }
        public void citacion4(Office.IRibbonControl control)
        {
            //Autenticidad
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Es imprescindible que tu texto sea original, que sean tus ideas y tus palabras. Pero asegúrate de incluir referencias precisas cada vez que utilices ideas, datos o palabras de otros autores.");
        }
        public void citacion5(Office.IRibbonControl control)
        {
            //Cita directa
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Cuando reproduces las ideas de otro debes aclarar quién es. Para ello es importante que escribas su apellido y el año de la publicación entre el paréntesis (En APA) o escribas el superíndice (en IEEE) al final.");
        }
        public void citacion6(Office.IRibbonControl control)
        {
            //Cita indirecta
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Cuando dentro de la idea que citas empleas el apellido del autor debes aclarar el año de la publicación entre paréntesis (en APA) o el superíndice (en IEEE).");
        }
        public void citacion7(Office.IRibbonControl control)
        {
            //Referencia de la cita en la lista
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Este texto no está enlistado en las referencias empleadas para hacer el trabajo. Debes incluirlo en la lista de referencias.");
        }
        public void citacion8(Office.IRibbonControl control)
        {
            //Referencia enlistada sin cita
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Este texto no hace parte del contenido. Revisa tu trabajo y si efectivamente lo usaste debes acomodar la cita de manera directa o indirecta y si no lo usaste sacar esta referencia. .");
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
            currentRange.Comments.Add(currentRange, "Para corroborar/respaldar esta afirmación debes citar y decir de dónde tomas la información. ");
        }
        public void citacion12(Office.IRibbonControl control)
        {
            // Cita posicionamiento
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "No olvides que en la cita de posicionamiento estás expresando tu opinión crítica sobre la fuente que estás citando. Así que asegúrate de que tu escritura refleje claramente tu punto de vista y cómo te distancias o te relacionas con el autor. También puedes considerar incluir una cita que respalde tu argumento y lo refuerce. Esto hará que tu posición sea más sólida y convincente.");
        }
        public void citacion13(Office.IRibbonControl control)
        {
            //Cita dialéctica
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que en la cita dialéctica comparas varios autores para ver cómo se relacionan sus ideas. Así que asegúrate de que estás presentando sus ideas de manera clara. Si solo estás utilizando una referencia, te sugiero buscar otra para enriquecer la conversación entre los autores. Esto hará que tu argumentación sea más sólida y fácil de entender.");
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
            currentRange.Comments.Add(currentRange, "Recuerda que el título principal debe ir centrado, en negrita, con cada letra inicial en mayúscula (excepto las preposiciones) y sin punto final.");
        }
        public void citacion16(Office.IRibbonControl control)
        {
            //subtítulos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que el subtítulo debe ir alineado a la izquierda, con cada letra inicial en mayúscula (excepto las preposiciones) y sin punto final. ");
        }
        public void citacion17(Office.IRibbonControl control)
        {
            //Notas al pie
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que las notas al pie se usan para expandir información de los autores del texto o la información contenida en él, no las uses para dar los datos de una referencia.");
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
            currentRange.Comments.Add(currentRange, "Recuerda que cuando se escriben textos muy largos es importante que crees  un índice donde queden claros los subtítulos y las páginas en las que el lector puede encontrarlos.");
        }
        public void citacion20(Office.IRibbonControl control)
        {
            //Mezcla de normas de citación
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Es importante recordar que al citar fuentes en un trabajo académico o científico, es necesario elegir y seguir una norma de citación específica. No se debes mezclar diferentes normas en un mismo trabajo. Cada norma tiene sus propias reglas y formatos, como IEEE con el uso de corchetes, Vancouver con paréntesis, Chicago con notas al pie de página, y APA que utiliza el apellido y el año de publicación entre paréntesis en el texto.");
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
            currentRange.Comments.Add(currentRange, " Usa signos de exclamación cuando emplees expresiones que denotan sorpresa, asombro, alegría, súplica, mandato, o deseo.");
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
            currentRange.Comments.Add(currentRange, "Recuerda que los puntos suspensivos debes usarlos al final de enumeraciones abiertas o incompletas, para reproducir una cita textual omitiendo una parte del final o para expresar duda. Van dentro de paréntesis (...) o corchetes [...] cuando al transcribir literalmente un texto se omite una parte de él.");
        }
        public void ED6(Office.IRibbonControl control)
        {
            //Sangría
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que para facilitar la comprensión de los párrafos debes sangría de primera línea en los párrafos sencillos, mientras que en las citas en bloque, los encabezados, las tablas y las figuras no se usa. ");
        }
        public void ED7(Office.IRibbonControl control)
        {
            //Sesgos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Es fundamental mostrar respeto por el lenguaje que las personas utilizan para describirse a sí mismas. Elegir etiquetas y términos con sensibilidad es esencial para garantizar el respeto hacia la individualidad y la humanidad de cada persona. Esto no solo implica utilizar un lenguaje inclusivo y respetuoso, sino también evitar perpetuar falsas jerarquías o estereotipos. La empatía y el respeto hacia las elecciones lingüísticas de los demás son componentes esenciales de una comunicación respetuosa y comprensiva.  ");
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
        public void ED15(Office.IRibbonControl control)
        {
            //Narrar
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Si tu objetivo con el texto es narrar, debes adecuarte a la estructura de dicho género discursivo para lograrlo.");
        }
        public void ED16(Office.IRibbonControl control)
        {
            //Describir
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Si tu objetivo con el texto es describir, debes adecuarte a la estructura de dicho género discursivo para lograrlo.");
        }
        public void ED17(Office.IRibbonControl control)
        {
            //Exponer
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Si tu objetivo con el texto es exponer sobre un tema, debes adecuarte a la estructura de dicho género discursivo para lograrlo.");
        }
        public void ED18(Office.IRibbonControl control)
        {
            //Argumentar
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Si tu objetivo con el texto es argumentar sobre un tema o una postura, debes adecuarte a la estructura de dicho género discursivo para lograrlo.");
        }
        public void ED19(Office.IRibbonControl control)
        {
            //Instruir
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Si tu objetivo con el texto es instruir o guiar, debes adecuarte a la estructura de dicho género discursivo para lograrlo.");
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
            currentRange.Comments.Add(currentRange, "Recuerda que para separar dos párrafos distintos que suelen desarrollar, dentro de la unidad del texto, contenidos diferentes, debe ir un punto aparte. Pero no lo uses al finalizar un título o un subtitulo, ni  en las indicaciones de lugar, ni fechas que encabezan cartas y documentos, ni en las cabeceras de tablas, ni después de signos de interrogación o exclamación.");
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
            currentRange.Comments.Add(currentRange, "Usa el punto y coma delante de las conjunciones adversativas (pero, más y aunque) cuando las oraciones que encabezan tienen cierta longitud.");
        }

        public void Delimitadores6(Office.IRibbonControl control)
        {
            //Dos puntos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que los dos puntos sirven para detener el discurso y llamar la atención sobre lo que sigue, se usan para anunciar enumeración, antes de citas textuales, para conectar oraciones y para presentar una conclusión o explicación. Entonces revisa este fragmento.");
        }
        public void Delimitadores7(Office.IRibbonControl control)
        {
            //Coma enumerativa
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Puedes usar la coma enumerativa para separar un conjunto de elementos que expresan características similares o simplemente se desean enumerar. Al nombrar estas palabras se usan conjunciones (y, o, u, ni), antes de ellas no se debe poner una coma. ");
        }
        public void Delimitadores8(Office.IRibbonControl control)
        {
            //Coma vocativa
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Puedes usar la coma vocativa para marcar una diferencia entre el vocativo y el resto de la oración. El vocativo es la forma de dirigirse a una persona, ya sea por su nombre o algo que la distinga.");
        }
        public void Delimitadores9(Office.IRibbonControl control)
        {
            //Coma elíptica
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Puedes usar la coma elíptica para sustituir un verbo o sustantivo que fue mencionado.");
        }
        public void Delimitadores10(Office.IRibbonControl control)
        {
            //Coma explicativa
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Puedes usar la coma explicativa o incidental para agregar datos adicionales del sujeto o el verbo.");
        }
        public void Delimitadores11(Office.IRibbonControl control)
        {
            //Coma apositiva
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, " Puedes usar la coma apositiva para agregar una explicación entre un nombre o grupo de nombres.  Ej.: María, mi vecina, regó las plantas en mis vacaciones.");
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
            currentRange.Comments.Add(currentRange, "Recuerda que en los textos académicos, debes emplear los guiones para unir palabras compuestas de manera temporal.");
        }
        public void Delimitadores15(Office.IRibbonControl control)
        {
            //Parentesis
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que cuando emplees el discurso directo debes usar los paréntesis para aclarar el apellido, la fecha de publicación de la cita y el número de página si es textual. Mientras que si se empleas el discurso indirecto en el paréntesis solo va la fecha de publicación y el número de página en caso de que sea textual. ");
        }
        public void Delimitadores16(Office.IRibbonControl control)
        {
            //Comillas
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que cuando uses una cita textual de menos de 40 palabras, la debes poner entre comillas. ");
        }
        public void Delimitadores17(Office.IRibbonControl control)
        {
            //Tilde
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Recuerda que debes tildar los monosílabos, las palabras agudas terminadas en -n o en -s que no tienen otra consonante antes de esa terminación, o las que terminen en las vocales a, e, i, o, u. También, las palabras llanas o graves que terminen en -y, una consonante que no sea -n o -s, o en más de una consonante. Y no olvides que todas las palabras esdrújulas y sobresdrújulas deben llevar tilde.");
        }
        public void Delimitadores18(Office.IRibbonControl control)
        {
            // Tilde diacrítica
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "¡Es muy importante recordar la tilde diacrítica! La necesitas para distinguir palabras que se escriben igual pero tienen significados diferentes. Generalmente, se utiliza en monosílabos como tu/tú, el/él, si/sí, dé/de, sé/se, o en palabras interrogativas y exclamativas como: cómo, cuándo, cuánto, (a)dónde, qué, cuál, cuán, quién... Usar la tilde adecuadamente en estas palabras es fundamental para comunicar con precisión. ¡Así evitaremos confusiones!");
        }
        public void Delimitadores19(Office.IRibbonControl control)
        {
            //Tilde en otras formas
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Los extranjerismos adaptados y las palabras compuestas siguen las reglas de acentuación prosódica, lo que significa que deben seguir las reglas de acentuación de las palabras graves, agudas y esdrújulas. Esto es importante para asegurarte de que estén escritos correctamente y se entienda su pronunciación.");
        }
        public void Delimitadores20(Office.IRibbonControl control)
        {
            //Mayúscula inicial
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Después de un punto seguido, un punto de interrogación o exclamación, es importante que utilices la mayúscula sostenida. Lo mismo aplica para los nombres propios de personas, cuerpos celestes, animales, plantas y objetos singularizados, signos del Zodiaco, accidentes geográficos, barrios, urbanizaciones, calles, espacios urbanos y vías de comunicación, establecimientos comerciales, culturales o recreativos, guerras, batallas, acontecimientos históricos relevantes, antropón");
        }
        public void Delimitadores21(Office.IRibbonControl control)
        {
            //Mayúscula sostenida
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Ten en cuenta que la mayúscula sostenida solo debes emplearla en cabeceras de diarios y revistas, en siglas y acrónimos, en números romanos y para hacer énfasis en algo fundamental del escrito y buscas llamar la atención del lector.");
        }
        public void Delimitadores22(Office.IRibbonControl control)
        {
            //Mayúscula después de dos puntos
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Después de dos puntos, solo debes poner mayúscula inicial cuando reproduces textualmente las palabras de otra persona, cuando estás haciendo una enumeración o cuando hay una completa independencia sintáctica y de sentido. ");
        }
        public void Delimitadores23(Office.IRibbonControl control)
        {
            //Mayúscula inicial de primera palabra
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Considera que debes poner mayúscula inicial a la primera palabra del título de cualquier obra de creación, a las subdivisiones o secciones internas de una publicación o un documento, a las palabras significativas que forman parte del nombre de eventos culturales o deportivos, de premios y condecoraciones.");
        }
        public void Delimitadores24(Office.IRibbonControl control)
        {
            //Minúscula
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Estima que debes escribir con minúscula los apodos, sobrenombres o seudónimos, los artículos que preceden  a los nombres de accidentes geográficos, los nombres de países, los continentes y los títulos abreviados de obras de creación. También debes usar minúscula en tratamientos, títulos nobiliarios, dignidades o cargos, profesiones, pueblos o etnias, nombres de las lenguas, puntos cardinales, hemisferios, las líneas imaginarias y los polos geográficos, días de la semana, meses y estaciones del año, las notas musicales, los principios activos de los medicamentos, las monedas, los nombres de las escuelas y corrientes de las diversas ramas del conocimiento, así como los de estilos, movimientos y géneros artísticos, las religiones, así como los sustantivos que designan el conjunto de sus fieles.");
        }
        public void Delimitadores25(Office.IRibbonControl control)
        {
            //Cursiva
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Ten en cuenta que debes poner cursiva en los extranjerismos no adaptados, incluidos los latinismos y nombres científicos. Igualmente, debes usarla en los títulos de las obras, tales como libros, discos o revistas, cuadros, películas, series de televisión...");
        }
        public void Delimitadores26(Office.IRibbonControl control)
        {
            //Redonda
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Considera que los nombres de emisoras, editoriales, casos policiales, leyes, colecciones de obras, las monedas y los nombres propios debes escribirlos con letra redonda.  ");
        }
        public void Delimitadores27(Office.IRibbonControl control)
        {
            //Coma criminal
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Revisa los usos de las comas porque estas no se usan para separar los sujetos de sus complementos ni para separar oraciones.");
        }
        public void Delimitadores28(Office.IRibbonControl control)
        {
            //Punto criminal
            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Comments.Add(currentRange, "Revisa el empleo del punto seguido, porque este se usa para separar dos enunciados con su sujeto, verbo y complemento independiente. Pero si no se has cambiado la acción o el sujeto del que hablas debes revisar si lo que necesitas es una coma o un punto y coma.");
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
