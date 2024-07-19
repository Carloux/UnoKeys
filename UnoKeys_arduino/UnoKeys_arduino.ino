const int buttonPins[] = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12};  // Pins to which the buttons are connected
const int numButtons = 11;  // Number of buttons
int buttonStates[numButtons];
int lastButtonStates[numButtons];

void setup() {
  Serial.begin(9600);
  for (int i = 0; i < numButtons; i++) {
    pinMode(buttonPins[i], INPUT_PULLUP);  
    buttonStates[i] = digitalRead(buttonPins[i]);
    lastButtonStates[i] = buttonStates[i];
  }
}

void loop() {
  for (int i = 0; i < numButtons; i++) {
    buttonStates[i] = digitalRead(buttonPins[i]);
    if (buttonStates[i] != lastButtonStates[i]) {
      if (buttonStates[i] == LOW) {  
        Serial.print("Button ");
        Serial.print(i);
        Serial.println(" Pressed");
      } else {
        Serial.print("Button ");
        Serial.print(i);
        Serial.println(" Released");
      }
      lastButtonStates[i] = buttonStates[i];
    }
  }
  delay(50);  
}
