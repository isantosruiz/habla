function habla(varargin)
%HABLA Lee cadenas de texto y números usando voz sintética
%
% habla(texto|número)
%
% Ejemplos:
%
%   habla('¡Hola!');
%   habla(-123.4);
%   habla('Números impares menores que 10:',1:2:9);
%   habla('El valor aproximado de pi es',pi);
%   expresion = '5 + 7';
%   habla('El resultado de',expresion,'es',eval(expresion))

try myExcel = actxserver('Excel.Application');
catch error('¡No se pudo acceder a Excel!'); end
texto='';
for k=1:nargin
    if ischar(varargin{k})
        texto = [texto ' ' varargin{k}];
    else
        texto = [texto ' ' num2str(varargin{k})];
    end
end
if isnumeric(texto)
    texto = num2str(texto);
end
try invoke(myExcel.Speech,'Speak',texto);
catch
    if ~exist('text','var')
        error('¡No has ingresado nada que hablar!');
    end
    invoke(myExcel,'Quit');
    delete(myExcel);
    error('¡La herramienta Speech no está disponible en Excel!');
end
invoke(myExcel,'Quit');
delete(myExcel);