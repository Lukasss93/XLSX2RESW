﻿//------------------------------------------------------------------------------
// <auto-generated>
//     Il codice è stato generato da uno strumento.
//     Versione runtime:4.0.30319.42000
//
//     Le modifiche apportate a questo file possono provocare un comportamento non corretto e andranno perse se
//     il codice viene rigenerato.
// </auto-generated>
//------------------------------------------------------------------------------

namespace XLSX2RESW.Properties {
    using System;
    
    
    /// <summary>
    ///   Classe di risorse fortemente tipizzata per la ricerca di stringhe localizzate e così via.
    /// </summary>
    // Questa classe è stata generata automaticamente dalla classe StronglyTypedResourceBuilder.
    // tramite uno strumento quale ResGen o Visual Studio.
    // Per aggiungere o rimuovere un membro, modificare il file con estensione ResX ed eseguire nuovamente ResGen
    // con l'opzione /str oppure ricompilare il progetto VS.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resources {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   Restituisce l'istanza di ResourceManager nella cache utilizzata da questa classe.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("XLSX2RESW.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Esegue l'override della proprietà CurrentUICulture del thread corrente per tutte le
        ///   ricerche di risorse eseguite utilizzando questa classe di risorse fortemente tipizzata.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Cerca una stringa localizzata simile a &lt;/root&gt;.
        /// </summary>
        internal static string resw_end {
            get {
                return ResourceManager.GetString("resw_end", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Cerca una stringa localizzata simile a &lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;
        ///&lt;root&gt;
        ///    &lt;xsd:schema id=&quot;root&quot; xmlns=&quot;&quot; xmlns:xsd=&quot;http://www.w3.org/2001/XMLSchema&quot; xmlns:msdata=&quot;urn:schemas-microsoft-com:xml-msdata&quot;&gt;
        ///        &lt;xsd:import namespace=&quot;http://www.w3.org/XML/1998/namespace&quot; /&gt;
        ///        &lt;xsd:element name=&quot;root&quot; msdata:IsDataSet=&quot;true&quot;&gt;
        ///            &lt;xsd:complexType&gt;
        ///                &lt;xsd:choice maxOccurs=&quot;unbounded&quot;&gt;
        ///                    &lt;xsd:element name=&quot;metadata&quot;&gt;
        ///                        &lt;xsd:complexType&gt;
        ///                      [stringa troncata]&quot;;.
        /// </summary>
        internal static string resw_start {
            get {
                return ResourceManager.GetString("resw_start", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Cerca una stringa localizzata simile a 	&lt;data name=&quot;{0}&quot; xml:space=&quot;preserve&quot;&gt;
        ///		&lt;value&gt;{1}&lt;/value&gt;
        ///	&lt;/data&gt;.
        /// </summary>
        internal static string resw_value {
            get {
                return ResourceManager.GetString("resw_value", resourceCulture);
            }
        }
    }
}
