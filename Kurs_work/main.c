#include "stm32f10x.h"
#include "diod.h"
#include <stdio.h>

#define COMMAND_BUFFER_SIZE 32

volatile uint32_t msTicks; 
volatile uint32_t temperature_update_flag = 1;
volatile uint32_t temperature_output_flag = 1;
volatile uint32_t micros_ticks = 0;
volatile uint32_t delay_micros_target = 0;
volatile uint32_t delay_micros_active = 0;

volatile uint8_t usart_rx_buffer[COMMAND_BUFFER_SIZE];
volatile uint8_t usart_rx_index = 0;
volatile uint8_t usart_command_ready = 0;

void SysTick_Handler(void) {
	static uint32_t measure_counter = 0;
  static uint32_t output_counter = 0;
  
	msTicks++;
	
	if(++measure_counter >= 1000) {
        temperature_update_flag = 1;
        measure_counter = 0;
    }
    
    if(++output_counter >= 10000) { 
        temperature_output_flag = 1;
        output_counter = 0;
    }
}

void DelayMicro (uint32_t dlyTicks) {
  uint32_t curTicks;

  curTicks = msTicks;
  while ((msTicks - curTicks) < dlyTicks) { __NOP(); }
}

void TIM2_IRQHandler(void) {
    if (TIM2->SR & TIM_SR_UIF) {
        TIM2->SR &= ~TIM_SR_UIF;
        
        if (delay_micros_active && micros_ticks >= delay_micros_target) {
            delay_micros_active = 0;
        }
    }
}

void Timer_Init(void) {
    RCC->APB1ENR |= RCC_APB1ENR_TIM2EN;
    TIM2->PSC = 72 - 1;   
    TIM2->ARR = 0xFFFF; 
    TIM2->EGR |= TIM_EGR_UG;
	
    TIM2->DIER |= TIM_DIER_UIE;
    NVIC_EnableIRQ(TIM2_IRQn);
    NVIC_SetPriority(TIM2_IRQn, 1);
    
    TIM2->CR1 |= TIM_CR1_CEN;
}

void SystemCoreClockConfigure(void) {
    RCC->CR |= ((uint32_t)RCC_CR_HSEBYP);                    
    while ((RCC->CR & RCC_CR_HSERDY) == 0);                  

    RCC->CFGR = RCC_CFGR_SW_HSE;                             
    while ((RCC->CFGR & RCC_CFGR_SWS) != RCC_CFGR_SWS_HSE);  
    
    RCC->CFGR = RCC_CFGR_HPRE_DIV1;                          // HCLK = SYSCLK
    RCC->CFGR |= RCC_CFGR_PPRE1_DIV1;                        // APB1 = HCLK/2
    RCC->CFGR |= RCC_CFGR_PPRE2_DIV1;                        // APB2 = HCLK/2

    RCC->CR &= ~RCC_CR_PLLON;                                

    // PLL configuration: = HSE * 9 = 72 MHz
    RCC->CFGR &= ~(RCC_CFGR_PLLSRC | RCC_CFGR_PLLMULL);
    RCC->CFGR |= (RCC_CFGR_PLLSRC_HSE | RCC_CFGR_PLLMULL9);

    RCC->CR |= RCC_CR_PLLON;                                
    while((RCC->CR & RCC_CR_PLLRDY) == 0) __NOP();           

    RCC->CFGR &= ~RCC_CFGR_SW;                               
    RCC->CFGR |= RCC_CFGR_SW_PLL;
    while ((RCC->CFGR & RCC_CFGR_SWS) != RCC_CFGR_SWS_PLL);  
}

void USART_Init () {
    RCC->APB1ENR |= RCC_APB1ENR_USART2EN; 
    RCC->APB2ENR |= RCC_APB2ENR_IOPAEN; 
    
    // PA2 (TX) - Alternate function push-pull
    GPIOA->CRL &= (~GPIO_CRL_CNF2_0); 
    GPIOA->CRL |= (GPIO_CRL_CNF2_1 | GPIO_CRL_MODE2);

    // PA3 (RX) - Input pull-up
    GPIOA->CRL &= (~GPIO_CRL_CNF3_0);
    GPIOA->CRL |= GPIO_CRL_CNF3_1;
    GPIOA->CRL &= (~(GPIO_CRL_MODE3));
    GPIOA->BSRR |= GPIO_ODR_ODR3;

    // USARTDIV = 72000000 / (16 * 9600) = 468.75
    USART2->BRR = 7500; // 0x1D4C

    USART2->CR1 |= USART_CR1_TE | USART_CR1_RE | USART_CR1_UE;
    USART2->CR2 = 0;
    USART2->CR3 = 0;
	
		USART2->CR1 |= USART_CR1_RXNEIE;
    
    NVIC_EnableIRQ(USART2_IRQn);
    NVIC_SetPriority(USART2_IRQn, 0);
}

void USART_Send (unsigned char symbol) {
    while ((USART2->SR & USART_SR_TC) == 0) {}
    USART2->DR = symbol;
}

void USART_Receive (unsigned char *data) {
    while ((USART2->SR & USART_SR_RXNE) == 0) {}
    *data = (unsigned char)USART2->DR;
}

void USART_SendString(char *str) {
	int count = 0;
  while (*str) {
		USART_Send(*str++);
		count++;
  }
}

void USART_SendUint8(uint8_t value) {
    if (value >= 100) {
        USART_Send((char)(48 + (value / 100) % 10));
    }
    if (value >= 10) {
        USART_Send((char)(48 + (value / 10) % 10));
    }
    USART_Send((char)(48 + value % 10));
}

void USART_SendFloat(float value) {
    int integer_part = (int)value;
    int decimal_part = (int)((double)(value - (float)integer_part) * 10000 + 0.5);
    
    if (integer_part < 0) {
        USART_Send('-');
        integer_part = -integer_part;
    }
    
    USART_SendUint8((uint8_t)value);
    
    USART_Send('.');
		USART_Send((char)(48 + (decimal_part / 1000) % 10));
		USART_Send((char)(48 + (decimal_part / 100) % 10));
    USART_Send((char)(48 + (decimal_part / 10) % 10));
    USART_Send((char)(48 + decimal_part % 10));
}

void SendTemperatureData(void) {
    uint8_t valid_sensors = 0;
	
    USART_SendString("Temperatures: ");
    for (int i = 0; i < MAX_SENSORS; i++) {
        if (!sensors[i].crc8_data_error && sensors[i].ROM_code[0] != 0) {
            valid_sensors++;
            
            if (valid_sensors > 1) {
                USART_SendString(" | ");
            }
            
            USART_Send('S');
            USART_Send((char)(48 + i));
            USART_SendString(": ");
            USART_SendFloat(sensors[i].temp);
            USART_Send('C');
        }
    }
    
    if (valid_sensors == 0) {
        USART_SendString("No sensors found");
    }
    
    USART_SendString("\r\n");
}

void USART2_IRQHandler(void) {
    if(USART2->SR & USART_SR_RXNE) {
        uint8_t received_char;
        USART_Receive(&received_char);
			
        USART_Send(received_char);
				USART_SendString("\r\n");
			
        
        if(received_char == '\r' || received_char == '\n') {
            if(usart_rx_index > 0) {
                usart_rx_buffer[usart_rx_index] = '\0';
                usart_command_ready = 1;
                usart_rx_index = 0;
            }
        } 
        else if(received_char >= 'a' && received_char <= 'h') {
            usart_rx_buffer[0] = received_char;
            usart_rx_buffer[1] = '\0';
            usart_command_ready = 1;
        }
        else if(usart_rx_index < COMMAND_BUFFER_SIZE - 1) {
            usart_rx_buffer[usart_rx_index++] = received_char;
        }
    }
}

void ChangeSensorResolution(uint8_t sensor_index, uint8_t sensor_bit_depth) {
		sensor_config[sensor_index].resolution = sensor_bit_depth;
		USART_SendString("Changed S");
		USART_Send((char)(48 + sensor_index));
		USART_SendString(" to ");
		USART_SendUint8(sensor_bit_depth);
		USART_SendString("-bit\r\n");
    
    sensor_config[sensor_index].changed = 1;
}

void ProcessCommand(uint8_t* command) {
    if(command[0] == 'a') {
        ChangeSensorResolution(0, 9);
    } 
    else if(command[0] == 'b') {
        ChangeSensorResolution(0, 10);
    }
		else if(command[0] == 'c') {
        ChangeSensorResolution(0, 11);
    }
		else if(command[0] == 'd') {
        ChangeSensorResolution(0, 12);
    }
		else if(command[0] == 'e') {
        ChangeSensorResolution(1, 9);
    }
		else if(command[0] == 'f') {
        ChangeSensorResolution(1, 10);
    }
		else if(command[0] == 'g') {
        ChangeSensorResolution(1, 11);
    }
		else if(command[0] == 'h') {
        ChangeSensorResolution(1, 12);
    }
}

uint8_t ReadSensorResolution(uint8_t sensor_index) {
    if (sensor_index >= MAX_SENSORS) {
        return 0;
    }
   
    uint8_t config_byte = sensors[sensor_index].scratchpad_data[4];
    uint8_t resolution_bits = (config_byte >> 5) & 0x03;
    
    switch(resolution_bits) {
        case 0x00:
            return 9;
        case 0x01:
            return 10;
        case 0x02:
            return 11;
        default:
            return 12;
    }
}

int main () {
    SystemCoreClockConfigure();
    SystemCoreClockUpdate();
    SysTick_Config(SystemCoreClock / 1000000);  // SysTick 1 msec interrupts
    
    Timer_Init();
    
    USART_Init();
    Init_Sensors();
	  
    diod();
    SendTemperatureData();
	
		ChangeSensorResolution(0, ReadSensorResolution(0));
		ChangeSensorResolution(1, ReadSensorResolution(1));
    
    while (1) {
        if (temperature_update_flag) {
            temperature_update_flag = 0;
            diod();
        }
        
        if (temperature_output_flag) {
						temperature_output_flag = 0;
            SendTemperatureData();
        }
				
				if(usart_command_ready) {
						usart_command_ready = 0;
						ProcessCommand((uint8_t*)usart_rx_buffer);
				}
    }
    
    return 0;
}