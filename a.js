let numeroProcesso1 = "8000866-24.2019.8.05.0216";
let numeroProcesso2 = "8057627-07.2020.8.05.0001";
let numeroProcesso3 = "8002271-56.2023.8.05.0216";

var partes1 = numeroProcesso1.split('.');
var partes2 = numeroProcesso2.split('.');
var partes3 = numeroProcesso3.split('.');

var ano1 = partes1[1];
var ano2 = partes2[1];
var ano3 = partes3[1];

console.log(ano1 ? ano1 : "Ano não encontrado"); // Deve retornar "2019"
console.log(ano2 ? ano2 : "Ano não encontrado"); // Deve retornar "2020"
console.log(ano3 ? ano3 : "Ano não encontrado"); // Deve retornar "2023"