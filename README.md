# MarCon
[ActiveX VB6] MarCon

Este foi um projeto complexo. Eu desenvolvi um dispositivo que era utilizado para marcar chapinhas de alumínio planas para identificação. Eu projetei uma placa eletrônica simples baseada em relês que eram controlados através dos pinos de uma porta RS232 de um computador. Também projetei e montei uma mesa de coordenadas com 2 eixos (X e Y), utilizando motores de passo que eram controlados por esta interface eletrônica, sendo que na mesa era fixada e chapinha que deveria ser marcada, e em um a posição fixa neste dispositivo foi instalado um pistão pneumático controlado através de uma válvula solenóide também ligada nesta interface.

Apesar deste dipositivo ser instalado em um único computador na rede, foi desenvolvido um controle ActiveX em Visual Basic que ficava dentro da intranet da empresa desenvolvida em PHP, onde era gerado as ordens de produção com os dados específicos do produto que eram passados por javascript ao controle ActiveX, dados estes que seriam marcados na chapinha.

------------
[ActiveX VB6] MarCon

This was a complex project. I developed a device that was used to mark flat aluminum plates for identification. I designed a simple electronic board based on relays that were controlled through the pins of a computer's RS232 port. I also designed and assembled a coordinate table with 2 axes (X and Y), using stepper motors that were controlled by this electronic interface, where the table was fixed and the plate that should be marked, and in one the fixed position in this device, a pneumatic piston controlled through a solenoid valve also connected to this interface was installed.

Despite this device being installed on a single computer on the network, an ActiveX control was developed in Visual Basic that was inside the company's intranet developed in PHP, where production orders were generated with the specific data of the product that were passed by javascript to the ActiveX control, data that would be marked on the flat plates.
